Module ExcelFunctions
    Public Sub UploadManufacturersToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка производителей в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportManufacturersHeaderToExcel(MyWRKBook, i)
        ExportManufacturersBodyToExcel(MyWRKBook, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadManufacturersToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка производителей в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportManufacturersHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        ExportManufacturersBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportManufacturersHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка производителей в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Список производителей для выгрузки на WEB сайт"
        MyWRKBook.ActiveSheet.Range("B1").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A3") = "Код поставщика"
        MyWRKBook.ActiveSheet.Range("B3") = "Имя поставщика"
        MyWRKBook.ActiveSheet.Range("C3") = "Имя поставщика для WEB"
        MyWRKBook.ActiveSheet.Range("D3") = "Резервное поле"

        MyWRKBook.ActiveSheet.Range("A3:D3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A3:D3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:D3").WrapText = True
        MyWRKBook.ActiveSheet.Range("A3:D3").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:D3").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:D3").Select()
        MyWRKBook.ActiveSheet.Range("A3:D3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:D3").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A3:D3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:D3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:D3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:D3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:D3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:D3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("C3:D3").Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 40

        i = 4
    End Function

    Public Function ExportManufacturersHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка производителей в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 4000
        oSheet.getColumns().getByName("B").Width = 8000
        oSheet.getColumns().getByName("C").Width = 8000
        oSheet.getColumns().getByName("D").Width = 8000

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список производителей для выгрузки на WEB сайт"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)

        i = 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код поставщика"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Имя поставщика"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Имя поставщика для WEB"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Резервное поле"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":D" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":D" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":D" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":D" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":D" & CStr(i))

        i = 4
    End Function

    Public Function ExportManufacturersBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела производителей в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT ID, Name, WEBName, Rezerv1 "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportManufacturersBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела производителей в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT ID, Name, WEBName, Rezerv1 "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(3)
                MyArr(0) = CDbl(Declarations.MyRec.Fields(0).Value)
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":D" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadManufacturersFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по производителям из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyCode As Double
        Dim MyWEBName As String
        Dim MyRezerv1 As String
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по производителям вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды производителей (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'C' и 'D' должны располагаться новые задаваемые вручную значения. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Название производителя для WEB и резервное поле (при необходимости)." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 4
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    Try
                        MyCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value
                        Try
                            If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                MyWEBName = ""
                            Else
                                MyWEBName = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value.ToString
                            End If
                            Try
                                If appXLSRC.Worksheets(1).Range("D" & StrCnt).Value = Nothing Then
                                    MyRezerv1 = ""
                                Else
                                    MyRezerv1 = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value.ToString
                                End If
                                Try
                                    MySQLStr = "UPDATE tbl_WEB_Manufacturers "
                                    MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(MyWEBName) & "', "
                                    MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(MyRezerv1) & "', "
                                    MySQLStr = MySQLStr & "RMStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' OR Rezerv1 <> N'" & Trim(MyRezerv1) & "' THEN 3 ELSE RMStatus END, "
                                    MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' OR Rezerv1 <> N'" & Trim(MyRezerv1) & "' THEN 3 ELSE WEBStatus END "
                                    MySQLStr = MySQLStr & "WHERE (ID = " & CStr(MyCode) & ") "
                                    MySQLStr = MySQLStr & "AND (RMStatus <> 2) "
                                    MySQLStr = MySQLStr & "AND (WEBStatus <> 2) "
                                    InitMyConn(False)
                                    Declarations.MyConn.Execute(MySQLStr)
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                            Catch ex As Exception
                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""D"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                            End Try
                        Catch ex As Exception
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        End Try
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                    End Try

                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadManufacturersFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по производителям из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim StrCnt As Double
        Dim MyCode As Double
        Dim MyWEBName As String
        Dim MyRezerv1 As String
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по производителям вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды производителей (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'C' и 'D' должны располагаться новые задаваемые вручную значения. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Название производителя для WEB и резервное поле (при необходимости)." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                StrCnt = 4
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '---код производителя
                    Try
                        MyCode = oSheet.getCellRangeByName("A" & StrCnt).value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    '---название производителя для WEB
                    MyWEBName = oSheet.getCellRangeByName("C" & StrCnt).String
                    '---поле rezerv
                    MyRezerv1 = oSheet.getCellRangeByName("D" & StrCnt).String

                    Try
                        MySQLStr = "UPDATE tbl_WEB_Manufacturers "
                        MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(MyWEBName) & "', "
                        MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(MyRezerv1) & "', "
                        MySQLStr = MySQLStr & "RMStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' OR Rezerv1 <> N'" & Trim(MyRezerv1) & "' THEN 3 ELSE RMStatus END, "
                        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' OR Rezerv1 <> N'" & Trim(MyRezerv1) & "' THEN 3 ELSE WEBStatus END "
                        MySQLStr = MySQLStr & "WHERE (ID = " & CStr(MyCode) & ") "
                        MySQLStr = MySQLStr & "AND (RMStatus <> 2) "
                        MySQLStr = MySQLStr & "AND (WEBStatus <> 2) "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadSalesmansToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продавцов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportSalesmansHeaderToExcel(MyWRKBook, i)
        ExportSalesmansBodyToExcel(MyWRKBook, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadSalesmansToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продавцов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportSalesmansHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        ExportSalesmansBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportSalesmansHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продавцов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Список продавцов для выгрузки на WEB сайт"
        MyWRKBook.ActiveSheet.Range("B1").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A3") = "Код продавца"
        MyWRKBook.ActiveSheet.Range("B3") = "Имя продавца"
        MyWRKBook.ActiveSheet.Range("C3") = "E-mail продавца"
        MyWRKBook.ActiveSheet.Range("D3") = "Город продавца (код)"
        MyWRKBook.ActiveSheet.Range("E3") = "Ответственный за WEB в городе"
        MyWRKBook.ActiveSheet.Range("F3") = "Дежурный"
        MyWRKBook.ActiveSheet.Range("G3") = "Активный"
        MyWRKBook.ActiveSheet.Range("H3") = "Резервное поле 1"
        MyWRKBook.ActiveSheet.Range("I3") = "Резервное поле 2"


        MyWRKBook.ActiveSheet.Range("A3:I3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A3:I3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:I3").WrapText = True
        MyWRKBook.ActiveSheet.Range("A3:I3").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:I3").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:I3").Select()
        MyWRKBook.ActiveSheet.Range("A3:I3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:I3").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A3:I3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:I3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:I3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:I3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:I3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:I3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("C3:I3").Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 40

        i = 4
    End Function


    Public Function ExportSalesmansHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка производителей в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 7600
        oSheet.getColumns().getByName("C").Width = 7600
        oSheet.getColumns().getByName("D").Width = 3800
        oSheet.getColumns().getByName("E").Width = 1900
        oSheet.getColumns().getByName("F").Width = 1900
        oSheet.getColumns().getByName("G").Width = 1900
        oSheet.getColumns().getByName("H").Width = 7600
        oSheet.getColumns().getByName("I").Width = 7600

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список продавцов для выгрузки на WEB сайт"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)

        i = 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код продавца"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Имя продавца"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "E-mail продавца"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Город продавца (код)"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Ответственный за WEB в городе"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Дежурный"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Активный"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "Резервное поле 1"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Резервное поле 2"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":I" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i))

        i = 4
    End Function


    Public Function ExportSalesmansBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела продавцов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT Code, Name, Email, ISNULL(City, 0) AS City, OfficeLeader, OnDuty, CONVERT(integer, IsActive) AS IsActive, Rezerv1, Rezerv2 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Salesmans "
        MySQLStr = MySQLStr & "ORDER BY Name "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportSalesmansBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела продавцов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT Code, Name, Email, ISNULL(City, 0) AS City, OfficeLeader, OnDuty, CONVERT(integer, IsActive) AS IsActive, Rezerv1, Rezerv2 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Salesmans "
        MySQLStr = MySQLStr & "ORDER BY Name "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(8)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = CInt(Declarations.MyRec.Fields(3).Value)
                MyArr(4) = CInt(Declarations.MyRec.Fields(4).Value)
                MyArr(5) = CInt(Declarations.MyRec.Fields(5).Value)
                MyArr(6) = CInt(Declarations.MyRec.Fields(6).Value)
                MyArr(7) = Declarations.MyRec.Fields(7).Value
                MyArr(8) = Declarations.MyRec.Fields(8).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadSalesmansFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продавцам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyCode As String
        Dim MyEmail As String
        Dim MyCity As Integer
        Dim MyWEBResp As Integer
        Dim MyOnDuty As Integer
        Dim MyActive As Integer
        Dim MyRez1 As String
        Dim MyRez2 As String
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по производителям вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продавцов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' адрес электронной почты. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' код города, в котором работает продавец (с закладки ""города""). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' является ли продавец ответственным за WEB в городе (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'F' является ли продавец дежурным (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'G' является ли продавец активным (выгружается ли на WEB сайт) (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'H' и 'I' - резервная информация (при необходимости)." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 4
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    MyCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value
                    If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                        MyEmail = ""
                    Else
                        MyEmail = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value.ToString
                    End If
                    Try
                        If appXLSRC.Worksheets(1).Range("D" & StrCnt).Value = Nothing Then
                            MyCity = 0
                        Else
                            MyCity = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                        End If
                        Try
                            If appXLSRC.Worksheets(1).Range("E" & StrCnt).Value = Nothing Then
                                MyWEBResp = 0
                            Else
                                MyWEBResp = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                            End If
                            Try
                                If appXLSRC.Worksheets(1).Range("F" & StrCnt).Value = Nothing Then
                                    MyOnDuty = 0
                                Else
                                    MyOnDuty = appXLSRC.Worksheets(1).Range("F" & StrCnt).Value
                                End If
                                Try
                                    If appXLSRC.Worksheets(1).Range("G" & StrCnt).Value = Nothing Then
                                        MyActive = 0
                                    Else
                                        MyActive = appXLSRC.Worksheets(1).Range("G" & StrCnt).Value
                                    End If
                                    Try
                                        If appXLSRC.Worksheets(1).Range("H" & StrCnt).Value = Nothing Then
                                            MyRez1 = ""
                                        Else
                                            MyRez1 = appXLSRC.Worksheets(1).Range("H" & StrCnt).Value.ToString
                                        End If
                                        Try
                                            If appXLSRC.Worksheets(1).Range("I" & StrCnt).Value = Nothing Then
                                                MyRez2 = ""
                                            Else
                                                MyRez2 = appXLSRC.Worksheets(1).Range("I" & StrCnt).Value.ToString
                                            End If
                                            Try
                                                MySQLStr = "UPDATE tbl_WEB_Salesmans "
                                                MySQLStr = MySQLStr & "SET Email = N'" & Trim(MyEmail) & "', "
                                                MySQLStr = MySQLStr & "City = " & CStr(MyCity) & ", "
                                                If MyWEBResp = 1 Then
                                                    MySQLStr = MySQLStr & "OfficeLeader = N'1', "
                                                Else
                                                    MySQLStr = MySQLStr & "OfficeLeader = N'0', "
                                                End If
                                                If MyOnDuty = 1 Then
                                                    MySQLStr = MySQLStr & "OnDuty = N'1', "
                                                Else
                                                    MySQLStr = MySQLStr & "OnDuty = N'0', "
                                                End If
                                                If MyActive = 1 Then
                                                    MySQLStr = MySQLStr & "IsActive = 1, "
                                                Else
                                                    MySQLStr = MySQLStr & "IsActive = 0, "
                                                End If
                                                MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(MyRez1) & "', "
                                                MySQLStr = MySQLStr & "Rezerv2 = N'" & Trim(MyRez2) & "' "
                                                If MyActive = 1 Then
                                                    MySQLStr = MySQLStr & ", RMStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE 3 END "
                                                    MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE 3 END "
                                                Else
                                                    MySQLStr = MySQLStr & ", RMStatus = CASE WHEN IsActive = 1 THEN 2 ELSE RMStatus END "
                                                    MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN IsActive = 1 THEN 2 ELSE WEBStatus END "
                                                End If
                                                MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                                                InitMyConn(False)
                                                Declarations.MyConn.Execute(MySQLStr)
                                            Catch ex As Exception
                                                MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        Catch ex As Exception
                                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""I"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                        End Try
                                    Catch ex As Exception
                                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""H"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                    End Try
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""G"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                            Catch ex As Exception
                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""F"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                            End Try
                        Catch ex As Exception
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""E"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        End Try
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""D"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                    End Try

                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadSalesmansFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продавцам из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim StrCnt As Double
        Dim MyCode As String
        Dim MyEmail As String
        Dim MyCity As Integer
        Dim MyWEBResp As Integer
        Dim MyOnDuty As Integer
        Dim MyActive As Integer
        Dim MyRez1 As String
        Dim MyRez2 As String
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по производителям вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продавцов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' адрес электронной почты. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' код города, в котором работает продавец (с закладки ""города""). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' является ли продавец ответственным за WEB в городе (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'F' является ли продавец дежурным (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'G' является ли продавец активным (выгружается ли на WEB сайт) (0 - нет, 1 - да). " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'H' и 'I' - резервная информация (при необходимости)." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                StrCnt = 4
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '---код продавца
                    MyCode = oSheet.getCellRangeByName("A" & CStr(StrCnt)).String
                    '---Email продавца
                    MyEmail = oSheet.getCellRangeByName("C" & CStr(StrCnt)).String
                    '---город продавца
                    Try
                        MyCity = oSheet.getCellRangeByName("D" & CStr(StrCnt)).value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""D"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Cities "
                    MySQLStr = MySQLStr & "WHERE (ID = " & MyCity & ") "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ячейка D" & CStr(StrCnt) & " Ошибка проверки кода города. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit While
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            trycloseMyRec()
                            MsgBox("Ячейка D" & CStr(StrCnt) & " Город с кодом " & CStr(MyCity) & " не найден.", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            Exit While
                        Else
                        End If
                    End If
                    '---ответственный за WEB
                    Try
                        MyWEBResp = oSheet.getCellRangeByName("E" & CStr(StrCnt)).value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""E"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyWEBResp <> 0 And MyWEBResp <> 1 Then
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""E"". Должно быть значение 0 или 1", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '---дежурный
                    Try
                        MyOnDuty = oSheet.getCellRangeByName("F" & CStr(StrCnt)).value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""F"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyOnDuty <> 0 And MyOnDuty <> 1 Then
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""F"". Должно быть значение 0 или 1", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '---активный
                    Try
                        MyActive = oSheet.getCellRangeByName("G" & CStr(StrCnt)).value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""G"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyActive <> 0 And MyActive <> 1 Then
                        MsgBox("Ошибка в строке " & CStr(StrCnt) & " колонке ""G"". Должно быть значение 0 или 1", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '---резервное поле 1
                    MyRez1 = oSheet.getCellRangeByName("H" & CStr(StrCnt)).value
                    '---резервное поле 2
                    MyRez2 = oSheet.getCellRangeByName("I" & CStr(StrCnt)).value

                    Try
                        MySQLStr = "UPDATE tbl_WEB_Salesmans "
                        MySQLStr = MySQLStr & "SET Email = N'" & Trim(MyEmail) & "', "
                        MySQLStr = MySQLStr & "City = " & CStr(MyCity) & ", "
                        MySQLStr = MySQLStr & "OfficeLeader = N'" & CStr(MyWEBResp) & "', "
                        MySQLStr = MySQLStr & "OnDuty = N'" & CStr(MyOnDuty) & "', "
                        MySQLStr = MySQLStr & "IsActive = " & CStr(MyActive) & ", "
                        MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(MyRez1) & "', "
                        MySQLStr = MySQLStr & "Rezerv2 = N'" & Trim(MyRez2) & "' "
                        If MyActive = 1 Then
                            MySQLStr = MySQLStr & ", RMStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE 3 END "
                            MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE 3 END "
                        Else
                            MySQLStr = MySQLStr & ", RMStatus = CASE WHEN IsActive = 1 THEN 2 ELSE RMStatus END "
                            MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN IsActive = 1 THEN 2 ELSE WEBStatus END "
                        End If
                        MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadProductGroupsToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка групп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportProductGroupsHeaderToExcel(MyWRKBook, i)
        ExportProductGroupsBodyToExcel(MyWRKBook, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadProductGroupsToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка групп продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportProductGroupsHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
                    oSheet, oFrame, i)
        ExportProductGroupsBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportProductGroupsHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка групп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Список групп продуктов для выгрузки на WEB сайт"
        MyWRKBook.ActiveSheet.Range("B1").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A3") = "Код группы"
        MyWRKBook.ActiveSheet.Range("B3") = "Имя группы"
        MyWRKBook.ActiveSheet.Range("C3") = "Имя группы для WEB"

        MyWRKBook.ActiveSheet.Range("A3:C3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A3:C3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:C3").WrapText = True
        MyWRKBook.ActiveSheet.Range("A3:C3").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:C3").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:C3").Select()
        MyWRKBook.ActiveSheet.Range("A3:C3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:C3").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A3:C3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:C3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:C3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:C3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:C3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:C3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("C3:C3").Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 60
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 60

        i = 4
    End Function

    Public Function ExportProductGroupsHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка групп продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 11400
        oSheet.getColumns().getByName("C").Width = 11400

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список групп продуктов для выгрузки на WEB сайт"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)

        i = 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код группы"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Имя группы"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Имя группы для WEB"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))

        i = 4
    End Function

    Public Function ExportProductGroupsBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела групп товаров в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT Code, Name, WEBName "
        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "ORDER BY Code "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportProductGroupsBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела групп товаров в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT Code, Name, WEBName "
        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "ORDER BY Code "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(2)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadProductGroupsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по группам продуктов из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyCode As String
        Dim MyWEBName As String
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по группам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться новые задаваемые вручную значения - " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Название группы продуктов для WEB." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 4
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    Try
                        MyCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value
                        Try
                            If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                MyWEBName = ""
                            Else
                                MyWEBName = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value.ToString
                            End If

                            Try
                                MySQLStr = "UPDATE tbl_WEB_ItemGroup "
                                MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(MyWEBName) & "', "
                                MySQLStr = MySQLStr & "RMStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' THEN 3 ELSE RMStatus END, "
                                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' THEN 3 ELSE WEBStatus END "
                                MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                                MySQLStr = MySQLStr & "AND (RMStatus <> 2) "
                                MySQLStr = MySQLStr & "AND (WEBStatus <> 2) "
                                InitMyConn(False)
                                Declarations.MyConn.Execute(MySQLStr)
                            Catch ex As Exception
                                MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                            End Try
                        Catch ex As Exception
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        End Try
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                    End Try

                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadProductGroupsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по группам продуктов из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim StrCnt As Double
        Dim MyCode As String
        Dim MyWEBName As String
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по группам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться новые задаваемые вручную значения - " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Название группы продуктов для WEB." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                StrCnt = 4
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    MyCode = oSheet.getCellRangeByName("A" & StrCnt).String
                    MyWEBName = oSheet.getCellRangeByName("C" & StrCnt).String
                    Try
                        MySQLStr = "UPDATE tbl_WEB_ItemGroup "
                        MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(MyWEBName) & "', "
                        MySQLStr = MySQLStr & "RMStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' THEN 3 ELSE RMStatus END, "
                        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBName <> N'" & Trim(MyWEBName) & "' THEN 3 ELSE WEBStatus END "
                        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
                        MySQLStr = MySQLStr & "AND (RMStatus <> 2) "
                        MySQLStr = MySQLStr & "AND (WEBStatus <> 2) "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadProductSubGroupsToExcel(ByVal MyGroup As String, ByVal MyGroupName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка подгрупп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportProductSubGroupsHeaderToExcel(MyWRKBook, i, MyGroup, MyGroupName)
        ExportProductSubGroupsBodyToExcel(MyWRKBook, i, MyGroup)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadProductSubGroupsToLO(ByVal MyGroup As String, ByVal MyGroupName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка подгрупп продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportProductSubGroupsHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyGroup, MyGroupName)
        ExportProductSubGroupsBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyGroup)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportProductSubGroupsHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyGroup As String, ByVal MyGroupName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка подгрупп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(MyGroup) = "" Then
            MyWRKBook.ActiveSheet.Range("B1") = "Список подгрупп продуктов для выгрузки на WEB сайт для всех групп"
        Else
            MyWRKBook.ActiveSheet.Range("B1") = "Список подгрупп продуктов для выгрузки на WEB сайт для группы " & MyGroup & " " & MyGroupName
        End If

        MyWRKBook.ActiveSheet.Range("B1").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A3") = "Код подгруппы"
        MyWRKBook.ActiveSheet.Range("B3") = "Код группы"
        MyWRKBook.ActiveSheet.Range("C3") = "Имя подгруппы"
        MyWRKBook.ActiveSheet.Range("D3") = "Описание подгруппы"
        MyWRKBook.ActiveSheet.Range("E3") = "Резервное поле"

        MyWRKBook.ActiveSheet.Range("A3:E3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A3:E3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:E3").WrapText = True
        MyWRKBook.ActiveSheet.Range("A3:E3").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:E3").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:E3").Select()
        MyWRKBook.ActiveSheet.Range("A3:E3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:E3").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A3:E3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:E3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:E3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:E3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:E3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:E3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("C3:E3").Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20

        i = 4
    End Function

    Public Function ExportProductSubGroupsHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyGroup As String, ByVal MyGroupName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка подгрупп продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 1900
        oSheet.getColumns().getByName("C").Width = 9500
        oSheet.getColumns().getByName("D").Width = 15200
        oSheet.getColumns().getByName("E").Width = 3800

        If Trim(MyGroup) = "" Then
            oSheet.getCellRangeByName("B" & CStr(i)).String = "Список подгрупп продуктов для выгрузки на WEB сайт для всех групп"
        Else
            oSheet.getCellRangeByName("B" & CStr(i)).String = "Список подгрупп продуктов для выгрузки на WEB сайт для группы " & MyGroup & " " & MyGroupName
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)

        i = 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код подгруппы"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Код группы"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Имя подгруппы"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Описание подгруппы"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Резервное поле"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":E" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))

        i = 4
    End Function

    Public Function ExportProductSubGroupsBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyGroup As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела подгрупп товаров в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT SubgroupCode, GroupCode, Name, Description, Rezerv1 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemSubGroup "
        If MyGroup = "" Then
        Else
            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & MyGroup & "') "
        End If
        MySQLStr = MySQLStr & "ORDER BY GroupCode,SubgroupCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportProductSubGroupsBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyGroup As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела подгрупп товаров в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT SubgroupCode, GroupCode, Name, Description, Rezerv1 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemSubGroup "
        If MyGroup = "" Then
        Else
            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & MyGroup & "') "
        End If
        MySQLStr = MySQLStr & "ORDER BY GroupCode,SubgroupCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(4)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadProductSubGroupsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по подгруппам продуктов из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MySubGroupCode As String
        Dim MyGroupCode As String
        Dim MyName As String
        Dim MyDescription As String
        Dim MyRezerv As String
        Dim StrCnt As Double
        Dim MySQLStr As String
        Dim MyNewSubGroupCodeD As Double
        Dim MyNewSubGroupCode As String
        Dim MyNewSubgroupID As String

        MyTxt = "Для импорта данных по подгруппам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды подгрупп) и 'B' (коды групп) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды подгрупп продуктов с предшествующими нулями. Если подгруппа новая - должна быть занесена 'N'. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'B' должны располагаться коды групп продуктов - Скальские с предшествующими нулями." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'С' указывается название подгруппы продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должно быть размещено описание подгруппы продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' указывается резервная информация по подгруппе (если надо) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 4
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    MySubGroupCode = appXLSRC.Worksheets(1).Range("A" & StrCnt).Value
                    If Trim(MySubGroupCode) <> "" Then
                        If appXLSRC.Worksheets(1).Range("B" & StrCnt).Value = Nothing Then
                            MyGroupCode = ""
                        Else
                            MyGroupCode = appXLSRC.Worksheets(1).Range("B" & StrCnt).Value
                        End If
                        If Trim(MyGroupCode) <> "" Then
                            If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                MyName = ""
                            Else
                                MyName = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                            End If
                            If Trim(MyName) <> "" Then
                                If appXLSRC.Worksheets(1).Range("D" & StrCnt).Value = Nothing Then
                                    MyDescription = ""
                                Else
                                    MyDescription = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                                End If
                                If appXLSRC.Worksheets(1).Range("E" & StrCnt).Value = Nothing Then
                                    MyRezerv = ""
                                Else
                                    MyRezerv = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                End If
                                '----------------Проверка, что группа продукта присутствует в БД
                                MySQLStr = "SELECT COUNT(Code) AS CC "
                                MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                                MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyGroupCode) & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указанный код группы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                    trycloseMyRec()
                                Else
                                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указанный код группы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                        trycloseMyRec()
                                    Else
                                        trycloseMyRec()
                                        If Trim(MySubGroupCode) = "N" Then
                                            '---------------------------Создание подгруппы
                                            '---получение нового кода
                                            MySQLStr = "SELECT CONVERT(numeric, ISNULL(MAX(SubgroupCode), 0)) AS CC "
                                            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                                            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Trim(MyGroupCode) & "')"
                                            InitMyConn(False)
                                            InitMyRec(False, MySQLStr)
                                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                MyNewSubGroupCodeD = 0
                                                trycloseMyRec()
                                            Else
                                                MyNewSubGroupCodeD = Declarations.MyRec.Fields("CC").Value
                                                trycloseMyRec()
                                            End If
                                            MyNewSubGroupCodeD = MyNewSubGroupCodeD + 1
                                            MyNewSubGroupCode = Right("0000" & CStr(MyNewSubGroupCodeD), 4)
                                            MyNewSubgroupID = Trim(MyGroupCode) & MyNewSubGroupCode
                                            '---Запись нового значения
                                            Try
                                                MySQLStr = "INSERT INTO tbl_WEB_ItemSubGroup "
                                                MySQLStr = MySQLStr & "(SubgroupID, SubgroupCode, GroupCode, Name, Description, Rezerv1, RMStatus, WEBStatus) "
                                                MySQLStr = MySQLStr & "VALUES (N'" & MyNewSubgroupID & "', "
                                                MySQLStr = MySQLStr & "N'" & MyNewSubGroupCode & "', "
                                                MySQLStr = MySQLStr & "N'" & MyGroupCode & "', "
                                                MySQLStr = MySQLStr & "N'" & MyName & "', "
                                                MySQLStr = MySQLStr & "N'" & MyDescription & "', "
                                                MySQLStr = MySQLStr & "N'" & MyRezerv & "', "
                                                MySQLStr = MySQLStr & "1, "
                                                MySQLStr = MySQLStr & "1) "
                                                InitMyConn(False)
                                                Declarations.MyConn.Execute(MySQLStr)
                                            Catch ex As Exception
                                                MsgBox("Ошибка в строке " & StrCnt & " " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        Else
                                            '---------------------------Редактирование подгруппы
                                            '----------------Проверка, что подгруппа продукта присутствует в БД
                                            MySQLStr = "SELECT COUNT(*) AS CC "
                                            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                                            MySQLStr = MySQLStr & "WHERE (SubgroupCode = N'" & MySubGroupCode & "') "
                                            MySQLStr = MySQLStr & "AND (GroupCode = N'" & MyGroupCode & "') "
                                            InitMyConn(False)
                                            InitMyRec(False, MySQLStr)
                                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Указанный код подгруппы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                                trycloseMyRec()
                                            Else
                                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Указанный код подгруппы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                                Else
                                                    '---Запись нового значения
                                                    MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
                                                    MySQLStr = MySQLStr & "SET Name = N'" & MyName & "', "
                                                    MySQLStr = MySQLStr & "Description = N'" & MyDescription & "', "
                                                    MySQLStr = MySQLStr & "Rezerv1 = N'" & MyRezerv & "', "
                                                    MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, "
                                                    MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END "
                                                    MySQLStr = MySQLStr & "WHERE (SubgroupCode = N'" & MySubGroupCode & "') "
                                                    MySQLStr = MySQLStr & "AND (GroupCode = N'" & MyGroupCode & "') "
                                                    InitMyConn(False)
                                                    Declarations.MyConn.Execute(MySQLStr)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Название подгруппы товара обязательно должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
                            End If
                        Else
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указание кода группы товара обязательно.", MsgBoxStyle.Critical, "Внимание!")
                        End If
                    Else
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". Указание кода подгруппы товара обязательно (или, если подгруппа новая, поставить N).", MsgBoxStyle.Critical, "Внимание!")
                    End If
                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadProductSubGroupsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по подгруппам продуктов из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim StrCnt As Double
        Dim MySubGroupCode As String
        Dim MyGroupCode As String
        Dim MyName As String
        Dim MyDescription As String
        Dim MyRezerv As String
        Dim MySQLStr As String
        Dim MyNewSubGroupCodeD As Double
        Dim MyNewSubgroupID As String

        MyTxt = "Для импорта данных по подгруппам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды подгрупп) и 'B' (коды групп) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды подгрупп продуктов с предшествующими нулями. Если подгруппа новая - должна быть занесена 'N'. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'B' должны располагаться коды групп продуктов - Скальские с предшествующими нулями." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'С' указывается название подгруппы продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должно быть размещено описание подгруппы продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' указывается резервная информация по подгруппе (если надо) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                StrCnt = 4
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '---код группы
                    MyGroupCode = oSheet.getCellRangeByName("B" & StrCnt).String
                    If MyGroupCode.Equals("") Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указание кода группы товара обязательно.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '---Проверка, что группа продукта присутствует в БД
                    MySQLStr = "SELECT COUNT(Code) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyGroupCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указанный код группы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""B"". Указанный код группы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If
                    '---код подгруппы
                    MySubGroupCode = oSheet.getCellRangeByName("A" & StrCnt).String
                    '---назване подгруппы
                    MyName = oSheet.getCellRangeByName("C" & StrCnt).String
                    If MyName.Equals("") Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Название подгруппы товара обязательно должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '---описание подгруппы
                    MyDescription = oSheet.getCellRangeByName("D" & StrCnt).String
                    '---резерв
                    MyRezerv = oSheet.getCellRangeByName("E" & StrCnt).String

                    '---занесение
                    If Trim(MySubGroupCode) = "N" Then
                        '---------------------------Создание подгруппы
                        '---получение нового кода
                        '---получение нового кода
                        MySQLStr = "SELECT CONVERT(numeric, ISNULL(MAX(SubgroupCode), 0)) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                        MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Trim(MyGroupCode) & "')"
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            MyNewSubGroupCodeD = 0
                            trycloseMyRec()
                        Else
                            MyNewSubGroupCodeD = Declarations.MyRec.Fields("CC").Value
                            trycloseMyRec()
                        End If
                        MyNewSubGroupCodeD = MyNewSubGroupCodeD + 1
                        MySubGroupCode = Right("0000" & CStr(MyNewSubGroupCodeD), 4)
                        MyNewSubgroupID = Trim(MyGroupCode) & MySubGroupCode
                        '---Запись нового значения
                        Try
                            MySQLStr = "INSERT INTO tbl_WEB_ItemSubGroup "
                            MySQLStr = MySQLStr & "(SubgroupID, SubgroupCode, GroupCode, Name, Description, Rezerv1, RMStatus, WEBStatus) "
                            MySQLStr = MySQLStr & "VALUES (N'" & MyNewSubgroupID & "', "
                            MySQLStr = MySQLStr & "N'" & MySubGroupCode & "', "
                            MySQLStr = MySQLStr & "N'" & MyGroupCode & "', "
                            MySQLStr = MySQLStr & "N'" & MyName & "', "
                            MySQLStr = MySQLStr & "N'" & MyDescription & "', "
                            MySQLStr = MySQLStr & "N'" & MyRezerv & "', "
                            MySQLStr = MySQLStr & "1, "
                            MySQLStr = MySQLStr & "1) "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        Catch ex As Exception
                            MsgBox("Ошибка в строке " & StrCnt & " " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            Exit Sub
                        End Try
                    Else
                        '---------------------------Редактирование подгруппы
                        '----------------Проверка, что подгруппа продукта присутствует в БД
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                        MySQLStr = MySQLStr & "WHERE (SubgroupCode = N'" & MySubGroupCode & "') "
                        MySQLStr = MySQLStr & "AND (GroupCode = N'" & MyGroupCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Указанный код подгруппы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Declarations.MyRec.Fields("CC").Value = 0 Then
                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Указанный код подгруппы товара не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                trycloseMyRec()
                                oWorkBook.Close(True)
                                Exit Sub
                            Else
                                '---Запись нового значения
                                Try
                                    MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
                                    MySQLStr = MySQLStr & "SET Name = N'" & MyName & "', "
                                    MySQLStr = MySQLStr & "Description = N'" & MyDescription & "', "
                                    MySQLStr = MySQLStr & "Rezerv1 = N'" & MyRezerv & "', "
                                    MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, "
                                    MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END "
                                    MySQLStr = MySQLStr & "WHERE (SubgroupCode = N'" & MySubGroupCode & "') "
                                    MySQLStr = MySQLStr & "AND (GroupCode = N'" & MyGroupCode & "') "
                                    InitMyConn(False)
                                    Declarations.MyConn.Execute(MySQLStr)
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                    trycloseMyRec()
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End Try
                            End If
                        End If
                    End If

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadProductsToExcel(ByVal MyGroup As String, ByVal MyGroupName As String, ByVal MySubGroup As String, ByVal MySubGroupName As String, ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel
        '// MySubgroupFlag (0 - выгрузка всех продуктов, не принадлежащих ни к одной группе
        '// 1 - выгрузка продуктов, принадлежащих к конкретной подгруппе - MySubGroup,
        '// 2 - выгрузка всех продуктов - как с подгруппами, так и без)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportProductsHeaderToExcel(MyWRKBook, i, MyGroup, MyGroupName, MySubGroup, MySubGroupName, MySubgroupFlag)
        ExportProductsBodyToExcel(MyWRKBook, i, MyGroup, MySubGroup, MySubgroupFlag)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadProductsToLO(ByVal MyGroup As String, ByVal MyGroupName As String, ByVal MySubGroup As String, ByVal MySubGroupName As String, ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в LibreOffice
        '// MySubgroupFlag (0 - выгрузка всех продуктов, не принадлежащих ни к одной группе
        '// 1 - выгрузка продуктов, принадлежащих к конкретной подгруппе - MySubGroup,
        '// 2 - выгрузка всех продуктов - как с подгруппами, так и без)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportProductsHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyGroup, MyGroupName, MySubGroup, MySubGroupName, MySubgroupFlag)
        ExportProductsBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyGroup, MySubGroup, MySubgroupFlag)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportProductsHeaderToExcel(ByRef MyWRKBook As Object, _
        ByRef i As Integer, _
        ByVal MyGroup As String, _
        ByVal MyGroupName As String, _
        ByVal MySubGroup As String, _
        ByVal MySubGroupName As String, _
        ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MySubgroupFlag = 2 Then      'выгрузка всех продуктов - как с подгруппами, так и без
            If Trim(MyGroup) = "" Then  'выгрузка всех продуктов для всех групп
                MyWRKBook.ActiveSheet.Range("B1") = "Список всех продуктов"
            Else                        'выгрузка всех продуктов для конкретной группы
                MyWRKBook.ActiveSheet.Range("B1") = "Список всех продуктов, входящих в группу " & MyGroup & " " & MyGroupName
            End If
        ElseIf MySubgroupFlag = 1 Then  'выгрузка всех продуктов для конкретной группы и подгруппы
            MyWRKBook.ActiveSheet.Range("B1") = "Список всех продуктов, входящих в группу " & MyGroup & " " & MyGroupName & " и подгруппу " & MySubGroup & " " & MySubGroupName
        Else                            'выгрузка продуктов, не включенных ни в одну подгруппу
            If Trim(MyGroup) = "" Then  'выгрузка продуктов, не включенных ни в одну подгруппу для всех групп
                MyWRKBook.ActiveSheet.Range("B1") = "Список всех продуктов, не включенных в подгруппы"
            Else                        'выгрузка продуктов, не включенных ни в одну подгруппу для конкретной группы
                MyWRKBook.ActiveSheet.Range("B1") = "Список всех продуктов, не включенных в подгруппы для группы " & MyGroup & " " & MyGroupName
            End If
        End If

        MyWRKBook.ActiveSheet.Range("B1").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A3") = "Код запаса"
        MyWRKBook.ActiveSheet.Range("B3") = "Имя запаса"
        MyWRKBook.ActiveSheet.Range("C3") = "Имя запаса для WEB"
        MyWRKBook.ActiveSheet.Range("D3") = "Код производителя"
        MyWRKBook.ActiveSheet.Range("E3") = "Имя производителя"
        MyWRKBook.ActiveSheet.Range("F3") = "Код страны"
        MyWRKBook.ActiveSheet.Range("G3") = "Страна"
        MyWRKBook.ActiveSheet.Range("H3") = "Код запаса производителя"
        MyWRKBook.ActiveSheet.Range("I3") = "Код группы"
        MyWRKBook.ActiveSheet.Range("J3") = "Название группы"
        MyWRKBook.ActiveSheet.Range("K3") = "Код подгруппы"
        MyWRKBook.ActiveSheet.Range("L3") = "Название подгруппы"
        MyWRKBook.ActiveSheet.Range("M3") = "Описание запаса"
        MyWRKBook.ActiveSheet.Range("N3") = "Признак складского ассортимента"
        MyWRKBook.ActiveSheet.Range("O3") = "Единица измерения"
        MyWRKBook.ActiveSheet.Range("P3") = "Резервное поле"
        MyWRKBook.ActiveSheet.Range("Q3") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("R3") = "Поставщик"
        MyWRKBook.ActiveSheet.Range("S3") = "Наличие картинки"

        MyWRKBook.ActiveSheet.Range("A3:S3").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A3:S3").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:S3").WrapText = True
        MyWRKBook.ActiveSheet.Range("A3:S3").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:S3").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A3:S3").Select()
        MyWRKBook.ActiveSheet.Range("A3:S3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:S3").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A3:S3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:S3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:S3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:S3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:S3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:S3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("C3;K3;M3;P3").Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 40

        i = 4
    End Function

    Public Function ExportProductsHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, _
        ByRef i As Integer, _
        ByVal MyGroup As String, _
        ByVal MyGroupName As String, _
        ByVal MySubGroup As String, _
        ByVal MySubGroupName As String, _
        ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 9500
        oSheet.getColumns().getByName("C").Width = 9500
        oSheet.getColumns().getByName("D").Width = 3800
        oSheet.getColumns().getByName("E").Width = 7600
        oSheet.getColumns().getByName("F").Width = 3800
        oSheet.getColumns().getByName("G").Width = 5700
        oSheet.getColumns().getByName("H").Width = 3800
        oSheet.getColumns().getByName("I").Width = 1900
        oSheet.getColumns().getByName("J").Width = 9500
        oSheet.getColumns().getByName("K").Width = 1900
        oSheet.getColumns().getByName("L").Width = 9500
        oSheet.getColumns().getByName("M").Width = 15200
        oSheet.getColumns().getByName("N").Width = 1900
        oSheet.getColumns().getByName("O").Width = 5700
        oSheet.getColumns().getByName("P").Width = 5700
        oSheet.getColumns().getByName("Q").Width = 3800
        oSheet.getColumns().getByName("R").Width = 7600
        oSheet.getColumns().getByName("S").Width = 7600
        oSheet.getColumns().getByName("T").Width = 3800
        oSheet.getColumns().getByName("U").Width = 3800
        oSheet.getColumns().getByName("V").Width = 3800
        oSheet.getColumns().getByName("W").Width = 3800
        oSheet.getColumns().getByName("X").Width = 3800

        If MySubgroupFlag = 2 Then      'выгрузка всех продуктов - как с подгруппами, так и без
            If Trim(MyGroup) = "" Then  'выгрузка всех продуктов для всех групп
                oSheet.getCellRangeByName("B" & CStr(i)).String = "Список всех продуктов"
            Else                        'выгрузка всех продуктов для конкретной группы
                oSheet.getCellRangeByName("B" & CStr(i)).String = "Список всех продуктов, входящих в группу " & MyGroup & " " & MyGroupName
            End If
        ElseIf MySubgroupFlag = 1 Then  'выгрузка всех продуктов для конкретной группы и подгруппы
            oSheet.getCellRangeByName("B" & CStr(i)).String = "Список всех продуктов, входящих в группу " & MyGroup & " " & MyGroupName & " и подгруппу " & MySubGroup & " " & MySubGroupName
        Else                            'выгрузка продуктов, не включенных ни в одну подгруппу
            If Trim(MyGroup) = "" Then  'выгрузка продуктов, не включенных ни в одну подгруппу для всех групп
                oSheet.getCellRangeByName("B" & CStr(i)).String = "Список всех продуктов, не включенных в подгруппы"
            Else                        'выгрузка продуктов, не включенных ни в одну подгруппу для конкретной группы
                oSheet.getCellRangeByName("B" & CStr(i)).String = "Список всех продуктов, не включенных в подгруппы для группы " & MyGroup & " " & MyGroupName
            End If
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i), 10)

        i = 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код запаса"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Имя запаса"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Имя запаса для WEB"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Код производителя"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Имя производителя"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Код страны"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Страна"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "Код запаса производителя"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Код группы"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "Название группы"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "Код подгруппы"
        oSheet.getCellRangeByName("L" & CStr(i)).String = "Название подгруппы"
        oSheet.getCellRangeByName("M" & CStr(i)).String = "Описание запаса"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "Признак складского ассортимента"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "Единица измерения"
        oSheet.getCellRangeByName("P" & CStr(i)).String = "Резервное поле"
        oSheet.getCellRangeByName("Q" & CStr(i)).String = "Код товара поставщика"
        oSheet.getCellRangeByName("R" & CStr(i)).String = "Поставщик"
        oSheet.getCellRangeByName("S" & CStr(i)).String = "Наличие картинки"
        oSheet.getCellRangeByName("T" & CStr(i)).String = "Наличие на складе"
        oSheet.getCellRangeByName("U" & CStr(i)).String = "Ожидаемый уход"
        oSheet.getCellRangeByName("V" & CStr(i)).String = "распределено"
        oSheet.getCellRangeByName("W" & CStr(i)).String = "Ожидаемый приход"
        oSheet.getCellRangeByName("X" & CStr(i)).String = "Прогноз наличия на складе"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":X" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":X" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":X" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("K" & CStr(i) & ":K" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("M" & CStr(i) & ":M" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("P" & CStr(i) & ":P" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":X" & CStr(i))

        i = 4
    End Function

    Public Function ExportProductsBodyToExcel(ByRef MyWRKBook As Object, _
        ByRef i As Integer, _
        ByVal MyGroup As String, _
        ByVal MySubGroup As String, _
        ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела групп товаров в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim j As Double

        If MySubgroupFlag = 2 Then      'выгрузка всех продуктов - как с подгруппами, так и без
            If Trim(MyGroup) = "" Then  'выгрузка всех продуктов для всех групп
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
                MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
                MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
                MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
                MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
                MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
                MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
            Else                        'выгрузка всех продуктов для конкретной группы
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
                MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
                MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
                MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
                MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
                MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
                MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
                MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
            End If
        ElseIf MySubgroupFlag = 1 Then  'выгрузка всех продуктов для конкретной группы и подгруппы
            MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
            MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
            MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
            MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
            MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
            MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
            MySQLStr = MySQLStr & "AND (tbl_WEB_Items.SubGroupCode = N'" & Trim(MySubGroup) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        Else                            'выгрузка продуктов, не включенных ни в одну подгруппу
            If Trim(MyGroup) = "" Then  'выгрузка продуктов, не включенных ни в одну подгруппу для всех групп
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
                MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
                MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
                MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
                MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
                MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
                MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
                MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.SubGroupCode = N'') "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
            Else                        'выгрузка продуктов, не включенных ни в одну подгруппу для конкретной группы
                MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
                MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
                MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
                MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
                MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
                MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
                MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
                MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
                MySQLStr = MySQLStr & "AND (tbl_WEB_Items.SubGroupCode = N'') "
                MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
            End If
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportProductsBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, _
        ByRef i As Integer, _
        ByVal MyGroup As String, _
        ByVal MySubGroup As String, _
        ByVal MySubgroupFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела групп товаров в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        'If MySubgroupFlag = 2 Then      'выгрузка всех продуктов - как с подгруппами, так и без
        '    If Trim(MyGroup) = "" Then  'выгрузка всех продуктов для всех групп
        '        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        '        MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        '        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
        '        MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
        '        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
        '        MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
        '        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        '        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
        '        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        '    Else                        'выгрузка всех продуктов для конкретной группы
        '        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        '        MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        '        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
        '        MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
        '        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
        '        MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
        '        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        '        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
        '        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
        '        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        '    End If
        'ElseIf MySubgroupFlag = 1 Then  'выгрузка всех продуктов для конкретной группы и подгруппы
        '    MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        '    MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        '    MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
        '    MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
        '    MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
        '    MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
        '    MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        '    MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
        '    MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
        '    MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
        '    MySQLStr = MySQLStr & "AND (tbl_WEB_Items.SubGroupCode = N'" & Trim(MySubGroup) & "') "
        '    MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        'Else                            'выгрузка продуктов, не включенных ни в одну подгруппу
        '    If Trim(MyGroup) = "" Then  'выгрузка продуктов, не включенных ни в одну подгруппу для всех групп
        '        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        '        MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        '        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
        '        MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
        '        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
        '        MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
        '        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        '        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
        '        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.SubGroupCode = N'') "
        '        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        '    Else                        'выгрузка продуктов, не включенных ни в одну подгруппу для конкретной группы
        '        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        '        MySQLStr = MySQLStr & "N'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        '        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_Items.SubGroupCode, ISNULL(tbl_WEB_ItemSubGroup.Name, N'') "
        '        MySQLStr = MySQLStr & "AS SubGroupName, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, "
        '        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppCode, PL010300.PL01002 AS SuppName, CASE WHEN tbl_WEB_Pictures.Picture IS NULL "
        '        MySQLStr = MySQLStr & "THEN '' ELSE '+' END AS Picture "
        '        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        '        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Items.Code = SC010300.SC01001 LEFT OUTER JOIN "
        '        MySQLStr = MySQLStr & "PL010300 ON PL010300.PL01001 = SC010300.SC01058 "
        '        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Trim(MyGroup) & "') "
        '        MySQLStr = MySQLStr & "AND (tbl_WEB_Items.SubGroupCode = N'') "
        '        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.GroupCode, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Code "
        '    End If
        'End If

        If MySubgroupFlag = 2 Then      'выгрузка всех продуктов - как с подгруппами, так и без
            If Trim(MyGroup) = "" Then  'выгрузка всех продуктов для всех групп
                MySQLStr = "exec dbo.spp_WEB_Items_FromDB_Extd N'', N'', 2 "
            Else                        'выгрузка всех продуктов для конкретной группы
                MySQLStr = "exec dbo.spp_WEB_Items_FromDB_Extd N'" & Trim(MyGroup) & "', N'', 2 "
            End If
        ElseIf MySubgroupFlag = 1 Then  'выгрузка всех продуктов для конкретной группы и подгруппы
            MySQLStr = "exec dbo.spp_WEB_Items_FromDB_Extd N'" & Trim(MyGroup) & "', N'" & Trim(MySubGroup) & "', 1 "
        Else                            'выгрузка продуктов, не включенных ни в одну подгруппу
            If Trim(MyGroup) = "" Then  'выгрузка продуктов, не включенных ни в одну подгруппу для всех групп
                MySQLStr = "exec dbo.spp_WEB_Items_FromDB_Extd N'', N'', 0 "
            Else                        'выгрузка продуктов, не включенных ни в одну подгруппу для конкретной группы
                MySQLStr = "exec dbo.spp_WEB_Items_FromDB_Extd N'" & Trim(MyGroup) & "', N'', 0 "
            End If
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(23)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = CInt(Declarations.MyRec.Fields(3).Value)
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArr(7) = Declarations.MyRec.Fields(7).Value
                MyArr(8) = Declarations.MyRec.Fields(8).Value
                MyArr(9) = Declarations.MyRec.Fields(9).Value
                MyArr(10) = Declarations.MyRec.Fields(10).Value
                MyArr(11) = Declarations.MyRec.Fields(11).Value
                MyArr(12) = Declarations.MyRec.Fields(12).Value
                MyArr(13) = CInt(Declarations.MyRec.Fields(13).Value)
                MyArr(14) = Declarations.MyRec.Fields(14).Value
                MyArr(15) = Declarations.MyRec.Fields(15).Value
                MyArr(16) = Declarations.MyRec.Fields(16).Value
                MyArr(17) = Declarations.MyRec.Fields(17).Value
                MyArr(18) = Declarations.MyRec.Fields(18).Value
                MyArr(19) = CDbl(Declarations.MyRec.Fields(19).Value)
                MyArr(20) = CDbl(Declarations.MyRec.Fields(20).Value)
                MyArr(21) = CDbl(Declarations.MyRec.Fields(21).Value)
                MyArr(22) = CDbl(Declarations.MyRec.Fields(22).Value)
                MyArr(23) = CDbl(Declarations.MyRec.Fields(23).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":X" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadProductsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продуктам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyScalaCode As String
        Dim MyWEBName As String
        Dim MyGroupCode As String
        Dim MySubGroupCode As String
        Dim MyDescription As String
        Dim MyRezerv As String
        Dim MySQLStr As String
        Dim StrCnt As String
        Dim MySubGroupFlag As Integer

        MyTxt = "Для импорта данных по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды товара в Scala) и 'I' (коды групп) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов Scala (с предшествующими нулями, если есть) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должно быть размещено описание продукта для выгрузки на WEB сайт " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'I' должны располагаться коды групп продуктов - Скальские с предшествующими нулями." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'K' указываются коды подгруппы продуктов - с предшествующими нулями " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'M' должно быть размещено описание продукта " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'P' указывается резервная информация по товару (если надо) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 4
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    MyScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                    If Trim(MyScalaCode) <> "" Then
                        If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                            MyWEBName = ""
                        Else
                            MyWEBName = Trim(appXLSRC.Worksheets(1).Range("C" & StrCnt).Value)
                        End If
                        If appXLSRC.Worksheets(1).Range("I" & StrCnt).Value = Nothing Then
                            MyGroupCode = ""
                        Else
                            MyGroupCode = Trim(appXLSRC.Worksheets(1).Range("I" & StrCnt).Value)
                        End If
                        If Trim(MyGroupCode) <> "" Then
                            If appXLSRC.Worksheets(1).Range("K" & StrCnt).Value = Nothing Then
                                MySubGroupCode = ""
                            Else
                                MySubGroupCode = Trim(appXLSRC.Worksheets(1).Range("K" & StrCnt).Value)
                            End If
                            If appXLSRC.Worksheets(1).Range("M" & StrCnt).Value = Nothing Then
                                MyDescription = ""
                            Else
                                MyDescription = Trim(appXLSRC.Worksheets(1).Range("M" & StrCnt).Value)
                            End If
                            If appXLSRC.Worksheets(1).Range("P" & StrCnt).Value = Nothing Then
                                MyRezerv = ""
                            Else
                                MyRezerv = Trim(appXLSRC.Worksheets(1).Range("P" & StrCnt).Value)
                            End If
                            '----------------Проверка, что скальский код с данной группой продуктов присутствует в БД
                            MySQLStr = "SELECT GroupCode "
                            MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
                            MySQLStr = MySQLStr & "WHERE (Code = N'" & MyScalaCode & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". Указанный код товара в Scala не найден в базе данных.", MsgBoxStyle.Critical, "Внимание!")
                                trycloseMyRec()
                            Else
                                If Trim(Declarations.MyRec.Fields("GroupCode").Value) <> MyGroupCode Then
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""I"". Указанный для данного товара код группы не соответствует тому, что проставлено для него в Scala - ." & Trim(Declarations.MyRec.Fields("GroupCode").Value) & ".", MsgBoxStyle.Critical, "Внимание!")
                                    trycloseMyRec()
                                Else
                                    trycloseMyRec()
                                    MySubGroupFlag = 0
                                    If MySubGroupCode = "" Then
                                        MySubGroupFlag = 0
                                    Else
                                        '--------Проверка, что указанный код подгруппы существует в данной группе
                                        MySQLStr = "SELECT COUNT(SubgroupCode) AS CC "
                                        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                                        MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & MyGroupCode & "') "
                                        MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & MySubGroupCode & "') "
                                        InitMyConn(False)
                                        InitMyRec(False, MySQLStr)
                                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""K"". Система не может проверить наличие подгруппы в базе данных. Обратитесь к администратору", MsgBoxStyle.Critical, "Внимание!")
                                            MySubGroupFlag = 1
                                            trycloseMyRec()
                                        Else
                                            If Trim(Declarations.MyRec.Fields("CC").Value) = 0 Then
                                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""K"". Указанный код подгруппы отсутствует в базе данных для указанного кода группы.", MsgBoxStyle.Critical, "Внимание!")
                                                MySubGroupFlag = 1
                                                trycloseMyRec()
                                            Else
                                                MySubGroupFlag = 0
                                                trycloseMyRec()
                                            End If
                                        End If
                                    End If

                                    If MySubGroupFlag = 0 Then
                                        Try
                                            '---Запись нового значения
                                            MySQLStr = "UPDATE tbl_WEB_Items "
                                            MySQLStr = MySQLStr & "SET WEBName = N'" & MyWEBName & "', "
                                            MySQLStr = MySQLStr & "SubGroupCode = N'" & MySubGroupCode & "', "
                                            MySQLStr = MySQLStr & "Description = N'" & MyDescription & "', "
                                            MySQLStr = MySQLStr & "Rezerv = N'" & MyRezerv & "', "
                                            MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE CASE WHEN RMStatus = 2 THEN 2 ELSE 3 END END, "
                                            MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE CASE WHEN WEBStatus = 2 THEN 2 ELSE 3 END END "
                                            MySQLStr = MySQLStr & "WHERE (Code = N'" & MyScalaCode & "')"
                                            InitMyConn(False)
                                            Declarations.MyConn.Execute(MySQLStr)
                                            '---Запись в историю
                                            MySQLStr = "DELETE FROM tbl_WEB_Items_InSubGroupHistory "
                                            MySQLStr = MySQLStr & "WHERE (Code = N'" & MyScalaCode & "') "
                                            InitMyConn(False)
                                            Declarations.MyConn.Execute(MySQLStr)

                                            If Trim(MySubGroupCode) <> "" Then
                                                MySQLStr = "INSERT INTO tbl_WEB_Items_InSubGroupHistory "
                                                MySQLStr = MySQLStr & "(Code, SubGroupCode) "
                                                MySQLStr = MySQLStr & "VALUES (N'" & MyScalaCode & "', "
                                                MySQLStr = MySQLStr & "N'" & Trim(MySubGroupCode) & "') "
                                                InitMyConn(False)
                                                Declarations.MyConn.Execute(MySQLStr)
                                            End If
                                        Catch ex As Exception
                                            MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                        End Try
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""I"". Указание кода группы товара обязательно.", MsgBoxStyle.Critical, "Внимание!")
                        End If
                    Else
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". Указание кода товара Scala обязательно.", MsgBoxStyle.Critical, "Внимание!")
                    End If
                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadProductsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продуктам из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim oCol As Object              '---Колонка, в которой смотрим диапазоны
        Dim oBlank As Object            '---набор пустых диапазонов
        Dim oRg                         '---пустой диапазон
        Dim oRange As Object
        Dim EndRange As Integer         '---Конец заполненного диапазона (строка начала пустого диапазона (ID))
        Dim StartRange As Integer
        Dim MySQLStr As String
        Dim MyTableName As String                   'Имя временной таблицы
        Dim MyGuid As String                          '
        Dim MySQLAdapter As SqlClient.SqlDataAdapter 'для временной таблицы
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyArr() As Object

        MyTxt = "Для импорта данных по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 4 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды товара в Scala) и 'I' (коды групп) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов Scala (с предшествующими нулями, если есть) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должно быть размещено описание продукта для выгрузки на WEB сайт " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'I' должны располагаться коды групп продуктов - Скальские с предшествующими нулями." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'K' указываются коды подгруппы продуктов - с предшествующими нулями " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'M' должно быть размещено описание продукта " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонках 'P' указывается резервная информация по товару (если надо) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '-----Вариант с импортом во временную таблицу
                '---определение диапазона данных
                StartRange = 4
                oCol = oSheet.Columns.getByIndex(0)
                oBlank = oCol.queryEmptyCells()
                oRg = oBlank.getByIndex(1)
                EndRange = oRg.RangeAddress.StartRow

                MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
                MyTableName = "tbl_ItemsParameters_Tmp_" + MyGuid
                '---Создание временных таблиц
                Try
                    MySQLStr = "DROP TABLE " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex As Exception
                End Try
                MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                MySQLStr = MySQLStr & "[Code] [nvarchar](35) NOT NULL, "
                MySQLStr = MySQLStr & "[WEBName] [nvarchar](250) NULL, "
                MySQLStr = MySQLStr & "[SubGroupCode] [nvarchar](50) NULL, "
                MySQLStr = MySQLStr & "[Description] [nvarchar](max) NULL, "
                MySQLStr = MySQLStr & "[Rezerv] [nvarchar](max) NULL "
                MySQLStr = MySQLStr & ") ON [PRIMARY] "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '-----данные 
                InitMyConn(False)
                MySQLStr = "SELECT [Code], [WEBName], [SubGroupCode], [Description], [Rezerv] "
                MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                Try
                    MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MySQLAdapter.SelectCommand.CommandTimeout = 1200
                    Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                    MySQLAdapter.Fill(MySQLDs)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex1 As Exception
                    End Try
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '-----Перенос данных из Excel dataset в SQL dataset
                Dim dt As DataTable
                Dim dr As DataRow

                dt = MySQLDs.Tables(0)
                oRange = oSheet.getCellRangeByName("A" & CStr(StartRange) & ":P" & CStr(EndRange))
                MyArr = oRange.DataArray
                For i As Integer = 0 To EndRange - 6
                    dr = dt.NewRow
                    '---код скала
                    If MyArr(i)(0).Equals("") Then
                        Exit For
                    End If
                    dr.Item(0) = MyArr(i)(0)
                    '---Имя запаса для WEB
                    dr.Item(1) = MyArr(i)(2)
                    '---код подгруппы
                    dr.Item(2) = MyArr(i)(10)
                    '---Описание запаса
                    dr.Item(3) = MyArr(i)(12)
                    '---резервное поле
                    dr.Item(4) = MyArr(i)(15)

                    dt.Rows.Add(dr)
                Next
                Try
                    MySQLAdapter.Update(MySQLDs, "Table")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex1 As Exception
                    End Try
                    oWorkBook.Close(True)
                    Exit Sub
                End Try

                '---проверка данных на сервере
                If ServerChecks(MyTableName) = False Then
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex1 As Exception
                    End Try
                    oWorkBook.Close(True)
                    Exit Sub
                End If

                '--------перенос данных в рабочую таблицу
                '---обновление значений
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET WEBName = " & MyTableName & ".WEBName, "
                MySQLStr = MySQLStr & "SubGroupCode = " & MyTableName & ".SubGroupCode, "
                MySQLStr = MySQLStr & "Description = " & MyTableName & ".Description, "
                MySQLStr = MySQLStr & "Rezerv = " & MyTableName & ".Rezerv, "
                MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE CASE WHEN RMStatus = 2 THEN 2 ELSE 3 END END, "
                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE CASE WHEN WEBStatus = 2 THEN 2 ELSE 3 END END "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
                MySQLStr = MySQLStr & MyTableName & " "
                MySQLStr = MySQLStr & "ON tbl_WEB_Items.Code = " & MyTableName & ".Code"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---удаление истории
                MySQLStr = "DELETE FROM tbl_WEB_Items_InSubGroupHistory "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Items_InSubGroupHistory INNER JOIN "
                MySQLStr = MySQLStr & MyTableName & " "
                MySQLStr = MySQLStr & "ON tbl_WEB_Items_InSubGroupHistory.Code = " & MyTableName & ".Code"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---занесение в историю
                MySQLStr = "INSERT INTO tbl_WEB_Items_InSubGroupHistory "
                MySQLStr = MySQLStr & "SELECT Code, SubGroupCode "
                MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                MySQLStr = MySQLStr & "WHERE (SubGroupCode <> N'') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---закрытие и удаление ненужного
                Try
                    MySQLStr = "DROP TABLE " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex1 As Exception
                End Try
                oWorkBook.Close(True)
            End If
        End If
    End Sub

    Private Function ServerChecks(ByVal MyTableName As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выполнение проверок данных на сервере  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLAdapter As SqlClient.SqlDataAdapter 'для результатов
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim WrkStr As String = ""
        Dim MySQLStr As String = ""
        Dim i As Integer

        MySQLStr = "spp_WEB_Items_UpdateFromExcel_Check N'" + Trim(MyTableName) + "'"
        InitMyConn(False)
        Try
            MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MySQLAdapter.SelectCommand.CommandTimeout = 1200
            Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
            MySQLAdapter.Fill(MySQLDs)
        Catch ex As Exception
            MsgBox(ex.ToString)
            ServerChecks = False
            Exit Function
        End Try

        '-----Заголовок
        If MySQLDs.Tables(0).Rows.Count > 0 Or MySQLDs.Tables(1).Rows.Count > 0 Then
            WrkStr = "Некорректно занесенная в Excel информация: " + Chr(13) + Chr(10)
        End If

        '-----коды товара
        If MySQLDs.Tables(0).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "Некорректные коды товаров: " + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(0).Rows.Count
                WrkStr = WrkStr + MySQLDs.Tables(0).Rows(i).Item(0) + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If

        '-----коды подгруппы товара
        If MySQLDs.Tables(1).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "Некорректные коды подгрупп товаров: " + Chr(13) + Chr(10)
            WrkStr = WrkStr + "Код товара                    Некорректный код подгруппы товара" + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(1).Rows.Count
                WrkStr = WrkStr + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(1).Rows(i).Item(0) + "                              ", 30) _
                    + MySQLDs.Tables(1).Rows(i).Item(2) + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If


        If WrkStr.Length > 0 Then
            MyErrorMessage = New ErrorMessage
            MyErrorMessage.TextBox1.Text = WrkStr
            MyErrorMessage.ShowDialog()
            ServerChecks = False
        Else
            ServerChecks = True
        End If
    End Function

    Public Sub UploadGroupDiscountToExcel(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на группы продуктов в Excel
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = MyCounter
        ExportGroupDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportGroupDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadGroupDiscountToLO(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на группы продуктов в LibreOffice
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = MyCounter
        '---скидки на группы
        ExportGroupDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportGroupDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportGroupDiscountHeaderToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на группы продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Список скидок на группы продуктов для клиента " & MyCustomerCodeName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Bold = True

        MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = MyCustomerCode
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Bold = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3)) = "Код группы товаров"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 3)) = "имя группы товаров"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 3)) = "Скидка (%)"

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ";C" & CStr(i + 3)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 15

        i = i + 4
    End Function

    Public Function ExportGroupDiscountHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на группы продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2750

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список скидок на группы продуктов для клиента " & MyCustomerCodeName
        oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), 10)

        oSheet.getCellRangeByName("C" & CStr(i)).String = MyCustomerCode
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))

        i = i + 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код группы товаров"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "имя группы товаров"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Скидка (%)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":A" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("B" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))

        i = i + 1
    End Function

    Public Function ExportGroupDiscountBodyToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на группы продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_DiscountGroup.GroupCode, ISNULL(tbl_WEB_ItemGroup.Name, N'') AS Name, tbl_WEB_DiscountGroup.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountGroup LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountGroup.GroupCode = tbl_WEB_ItemGroup.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountGroup.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountGroup.GroupCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            i = i + Declarations.MyRec.RecordCount
            trycloseMyRec()
        End If
    End Function

    Public Function ExportGroupDiscountBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на группы продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_WEB_DiscountGroup.GroupCode, ISNULL(tbl_WEB_ItemGroup.Name, N'') AS Name, tbl_WEB_DiscountGroup.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountGroup LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountGroup.GroupCode = tbl_WEB_ItemGroup.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountGroup.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountGroup.GroupCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(2)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CDbl(Declarations.MyRec.Fields(2).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadGroupDiscountsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по группам продуктов из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyDiscount As Double
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по скидкам по группам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по группам продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                '---------------Чтение и проверка кода клиента
                If appXLSRC.Worksheets(1).Range("C1").Value = Nothing Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    MyClientCode = Trim(appXLSRC.Worksheets(1).Range("C1").Value)
                    '---проверяем - есть ли такой клиент в БД, работающий через WEB
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                    MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            '---Очищаем старые значения из таблицы tbl_WEB_DiscountGroup
                            MySQLStr = "DELETE FROM tbl_WEB_DiscountGroup "
                            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)

                            '---ну и собственно импорт новых значений
                            StrCnt = 5
                            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                                Try
                                    MyCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                                    '---проверяем - есть ли такая группа в БД
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                        trycloseMyRec()
                                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такая группа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                    Else
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Группа с таким кодом отсутствует в базе. Импорт скидок по такой группе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                            trycloseMyRec()
                                        Else
                                            trycloseMyRec()
                                            Try
                                                If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                                    MyDiscount = 0
                                                Else
                                                    MyDiscount = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                                                End If
                                                If MyDiscount <= 0 Then
                                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                                                Else
                                                    Try
                                                        MySQLStr = "INSERT INTO tbl_WEB_DiscountGroup "
                                                        MySQLStr = MySQLStr & "(ID, ClientCode, GroupCode, Discount) "
                                                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                                                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                                                        MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ") "

                                                        InitMyConn(False)
                                                        Declarations.MyConn.Execute(MySQLStr)
                                                    Catch ex As Exception
                                                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                    End Try
                                                End If
                                            Catch ex As Exception
                                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                StrCnt = StrCnt + 1
                            End While
                        End If
                    End If
                End If

                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadGroupDiscountsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по группам продуктов из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyDiscount As Double
        Dim MySQLStr As String
        Dim StrCnt As Double

        MyTxt = "Для импорта данных по скидкам по группам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по группам продуктов " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '---------------Чтение и проверка кода клиента
                MyClientCode = oSheet.getCellRangeByName("C1").String
                If MyClientCode.Equals("") Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                End If
                '---проверяем - есть ли такой клиент в БД, работающий через WEB
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If
                End If

                '---Очищаем старые значения из таблицы tbl_WEB_DiscountGroup
                MySQLStr = "DELETE FROM tbl_WEB_DiscountGroup "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---ну и собственно импорт новых значений
                StrCnt = 5
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '-----код группы товаров
                    MyCode = Trim(oSheet.getCellRangeByName("A" & StrCnt).String)
                    '---проверяем - есть ли такая группа в БД
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такая группа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Группа с таким кодом отсутствует в базе. Импорт скидок по такой группе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            trycloseMyRec()
                        End If
                    End If
                    '-----Скидка
                    Try
                        MyDiscount = oSheet.getCellRangeByName("C" & StrCnt).Value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть числом" & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyDiscount <= 0 Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----занесение информации
                    Try
                        MySQLStr = "INSERT INTO tbl_WEB_DiscountGroup "
                        MySQLStr = MySQLStr & "(ID, ClientCode, GroupCode, Discount) "
                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                        MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ") "

                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadSubgroupDiscountToExcel(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на подгруппы продуктов в Excel
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = MyCounter
        ExportSubgroupDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportSubgroupDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadSubgroupDiscountToLO(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на подгруппы продуктов в LibreOffice
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = MyCounter
        '---скидки на подгруппы
        ExportSubgroupDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportSubgroupDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportSubgroupDiscountHeaderToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на подгруппы продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Список скидок на подгруппы продуктов для клиента " & MyCustomerCodeName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Bold = True

        MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = MyCustomerCode
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Bold = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3)) = "Код группы товаров"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 3)) = "имя группы товаров"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 3)) = "Код подгруппы товаров"
        MyWRKBook.ActiveSheet.Range("D" & CStr(i + 3)) = "имя подгруппы товаров"
        MyWRKBook.ActiveSheet.Range("E" & CStr(i + 3)) = "Скидка (%)"

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ";C" & CStr(i + 3) & ";E" & CStr(i + 3)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15

        i = i + 4
    End Function

    Public Function ExportSubgroupDiscountHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на подгруппы продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2750
        oSheet.getColumns().getByName("D").Width = 15200
        oSheet.getColumns().getByName("E").Width = 2750

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список скидок на подгруппы продуктов для клиента " & MyCustomerCodeName
        oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), 10)

        oSheet.getCellRangeByName("C" & CStr(i)).String = MyCustomerCode
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))

        i = i + 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код группы товаров"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "имя группы товаров"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Код подгруппы товаров"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "имя подгруппы товаров"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Скидка (%)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":A" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("B" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("D" & CStr(i) & ":D" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("E" & CStr(i) & ":E" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))

        i = i + 1
    End Function

    Public Function ExportSubgroupDiscountBodyToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на подгруппы продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_DiscountSubgroup.GroupCode, ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_DiscountSubgroup.SubgroupCode, "
        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemSubGroup.Name, N'') AS SubgroupName, tbl_WEB_DiscountSubgroup.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountSubgroup LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        MySQLStr = MySQLStr & "tbl_WEB_DiscountSubgroup.SubgroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemGroup.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountSubgroup.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountSubgroup.GroupCode, tbl_WEB_DiscountSubgroup.SubgroupCode "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            i = i + Declarations.MyRec.RecordCount
            trycloseMyRec()
        End If
    End Function

    Public Function ExportSubgroupDiscountBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на подгруппы продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_WEB_DiscountSubgroup.GroupCode, ISNULL(tbl_WEB_ItemGroup.Name, N'') AS GroupName, tbl_WEB_DiscountSubgroup.SubgroupCode, "
        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemSubGroup.Name, N'') AS SubgroupName, tbl_WEB_DiscountSubgroup.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountSubgroup LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        MySQLStr = MySQLStr & "tbl_WEB_DiscountSubgroup.SubgroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemGroup.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountSubgroup.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountSubgroup.GroupCode, tbl_WEB_DiscountSubgroup.SubgroupCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(4)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = CDbl(Declarations.MyRec.Fields(4).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadSubgroupDiscountsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по подгруппам продуктов из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MySubCode As String
        Dim MyDiscount As Double
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по скидкам по подгруппам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться коды подгрупп продуктов. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по подгруппам " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                '---------------Чтение и проверка кода клиента
                If appXLSRC.Worksheets(1).Range("C1").Value = Nothing Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    MyClientCode = Trim(appXLSRC.Worksheets(1).Range("C1").Value)
                    '---проверяем - есть ли такой клиент в БД, работающий через WEB
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                    MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            '---Очищаем старые значения из таблицы tbl_WEB_DiscountSubgroup
                            MySQLStr = "DELETE FROM tbl_WEB_DiscountSubgroup "
                            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)

                            '---ну и собственно импорт новых значений
                            StrCnt = 5
                            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                                Try
                                    MyCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                                    '---проверяем - есть ли такая группа в БД
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                        trycloseMyRec()
                                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такая группа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                        trycloseMyRec()
                                    Else
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Группа с таким кодом отсутствует в базе. Импорт скидок по такой группе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                            trycloseMyRec()
                                        Else
                                            trycloseMyRec()
                                            If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                                MySubCode = ""
                                            Else
                                                MySubCode = Trim(appXLSRC.Worksheets(1).Range("C" & StrCnt).Value)
                                            End If
                                            '---проверяем - есть ли такая подгруппа в БД
                                            MySQLStr = "SELECT COUNT(*) AS CC "
                                            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                                            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & MyCode & "') "
                                            MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & MySubCode & "') "
                                            InitMyConn(False)
                                            InitMyRec(False, MySQLStr)
                                            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                                trycloseMyRec()
                                                MsgBox("Ошибка в ячейке C" & StrCnt & ". Невозможно проверить, есть ли такая подгруппа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                            Else
                                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                                    MsgBox("Ошибка в ячейке C" & StrCnt & ". Подруппа с таким кодом отсутствует в базе. Импорт скидок по такой подгруппе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                                    trycloseMyRec()
                                                Else
                                                    trycloseMyRec()
                                                    Try
                                                        If appXLSRC.Worksheets(1).Range("E" & StrCnt).Value = Nothing Then
                                                            MyDiscount = 0
                                                        Else
                                                            MyDiscount = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value
                                                        End If
                                                        If MyDiscount <= 0 Then
                                                            MsgBox("Ошибка в строке " & StrCnt & " колонке ""E"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                                                        Else
                                                            Try
                                                                MySQLStr = "INSERT INTO tbl_WEB_DiscountSubgroup "
                                                                MySQLStr = MySQLStr & "(ID, ClientCode, GroupCode, SubgroupCode, Discount) "
                                                                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                                MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                                                                MySQLStr = MySQLStr & "N'" & MyCode & "', "
                                                                MySQLStr = MySQLStr & "N'" & MySubCode & "', "
                                                                MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ") "

                                                                InitMyConn(False)
                                                                Declarations.MyConn.Execute(MySQLStr)
                                                            Catch ex As Exception
                                                                MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                            End Try
                                                        End If
                                                    Catch ex As Exception
                                                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                    End Try
                                                End If
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                StrCnt = StrCnt + 1
                            End While
                        End If
                    End If
                End If

                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadSubgroupDiscountsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по подгруппам продуктов из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MySQLStr As String
        Dim StrCnt As Double
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MySubCode As String
        Dim MyDiscount As Double

        MyTxt = "Для импорта данных по скидкам по подгруппам продуктов вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды групп продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться коды подгрупп продуктов. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по подгруппам " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '---------------Чтение и проверка кода клиента
                MyClientCode = oSheet.getCellRangeByName("C1").String
                If MyClientCode.Equals("") Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                End If
                '---проверяем - есть ли такой клиент в БД, работающий через WEB
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If
                End If

                '---Очищаем старые значения из таблицы tbl_WEB_DiscountSubgroup
                MySQLStr = "DELETE FROM tbl_WEB_DiscountSubgroup "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---ну и собственно импорт новых значений
                StrCnt = 5
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '-----код группы товаров
                    MyCode = Trim(oSheet.getCellRangeByName("A" & StrCnt).String)
                    '---проверяем - есть ли такая группа в БД
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такая группа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Группа с таким кодом отсутствует в базе. Импорт скидок по такой группе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            trycloseMyRec()
                        End If
                    End If
                    '-----код подгруппы товаров
                    MySubCode = oSheet.getCellRangeByName("C" & StrCnt).String
                    '---проверяем - есть ли такая подгруппа в БД
                    If Not MySubCode.Equals("") Then
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                        MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & MyCode & "') "
                        MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & MySubCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            trycloseMyRec()
                            MsgBox("Ошибка в ячейке C" & StrCnt & ". Невозможно проверить, есть ли такая подгруппа в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Declarations.MyRec.Fields("CC").Value = 0 Then
                                MsgBox("Ошибка в ячейке C" & StrCnt & ". Подруппа с таким кодом отсутствует в базе. Импорт скидок по такой подгруппе невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                trycloseMyRec()
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    End If
                    '-----Скидка
                    Try
                        MyDiscount = oSheet.getCellRangeByName("E" & StrCnt).Value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""E"". Скидка должна быть числом" & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyDiscount <= 0 Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""E"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----занесение информации
                    Try
                        MySQLStr = "INSERT INTO tbl_WEB_DiscountSubgroup "
                        MySQLStr = MySQLStr & "(ID, ClientCode, GroupCode, SubgroupCode, Discount) "
                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                        MySQLStr = MySQLStr & "N'" & MySubCode & "', "
                        MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ") "

                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadItemDiscountToExcel(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на продукты в Excel
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = MyCounter
        ExportItemDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportItemDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadItemDiscountToLO(ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка скидок на продукты в LibreOffice
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = MyCounter
        '---скидки на отдельные товары
        ExportItemDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportItemDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportItemDiscountHeaderToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на продукты в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Список скидок на продукты для клиента " & MyCustomerCodeName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Bold = True

        MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = MyCustomerCode
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Bold = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3)) = "Код товара Scala"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 3)) = "Название товара"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 3)) = "Скидка (%)"

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":C" & CStr(i + 3)).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ";C" & CStr(i + 3)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 15

        i = i + 4
    End Function

    Public Function ExportItemDiscountHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, ByVal MyPriceInfo As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка скидок на продукты в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2750

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список скидок на продукты для клиента " & MyCustomerCodeName
        oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), 10)

        oSheet.getCellRangeByName("C" & CStr(i)).String = MyCustomerCode
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))

        i = i + 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код товара Scala"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Название товара"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Скидка (%)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":A" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("B" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))

        i = i + 1
    End Function

    Public Function ExportItemDiscountBodyToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на продукты в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_DiscountItem.ItemCode, tbl_WEB_Items.Name, tbl_WEB_DiscountItem.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountItem LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_DiscountItem.ItemCode = tbl_WEB_Items.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountItem.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountItem.ItemCode "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            i = i + Declarations.MyRec.RecordCount
            trycloseMyRec()
        End If
    End Function

    Public Function ExportItemDiscountBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела скидок на продукты в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_WEB_DiscountItem.ItemCode, tbl_WEB_Items.Name, tbl_WEB_DiscountItem.Discount "
        MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountItem LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_DiscountItem.ItemCode = tbl_WEB_Items.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountItem.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountItem.ItemCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(2)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CDbl(Declarations.MyRec.Fields(2).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadItemDiscountsFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по продуктам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyDiscount As Double
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по скидкам по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по товарам " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                '---------------Чтение и проверка кода клиента
                If appXLSRC.Worksheets(1).Range("C1").Value = Nothing Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    MyClientCode = Trim(appXLSRC.Worksheets(1).Range("C1").Value)
                    '---проверяем - есть ли такой клиент в БД, работающий через WEB
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                    MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            '---Очищаем старые значения из таблицы tbl_WEB_DiscountItem
                            MySQLStr = "DELETE FROM tbl_WEB_DiscountItem "
                            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)

                            '---ну и собственно импорт новых значений
                            StrCnt = 5
                            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                                Try
                                    MyCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                                    '---проверяем - есть ли такой товар в БД
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
                                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                        trycloseMyRec()
                                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такой товар в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                        trycloseMyRec()
                                    Else
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Товар с таким кодом отсутствует в базе. Импорт скидок по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                            trycloseMyRec()
                                        Else
                                            trycloseMyRec()
                                            Try
                                                If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                                    MyDiscount = 0
                                                Else
                                                    MyDiscount = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                                                End If
                                                If MyDiscount <= 0 Then
                                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                                                Else
                                                    Try
                                                        MySQLStr = "INSERT INTO tbl_WEB_DiscountItem "
                                                        MySQLStr = MySQLStr & "(ID, ItemCode, ClientCode, Discount) "
                                                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                                                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                                                        MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ")"

                                                        InitMyConn(False)
                                                        Declarations.MyConn.Execute(MySQLStr)
                                                    Catch ex As Exception
                                                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                    End Try
                                                End If
                                            Catch ex As Exception
                                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                StrCnt = StrCnt + 1
                            End While
                        End If
                    End If
                End If

                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadItemDiscountsFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по скидкам по продуктам из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MySQLStr As String
        Dim StrCnt As Double
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyDiscount As Double

        MyTxt = "Для импорта данных по скидкам по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружаются скидки." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения скидки в % " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по скидкам по товарам " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '---------------Чтение и проверка кода клиента
                MyClientCode = oSheet.getCellRangeByName("C1").String
                If MyClientCode.Equals("") Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                End If
                '---проверяем - есть ли такой клиент в БД, работающий через WEB
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If
                End If

                '---Очищаем старые значения из таблицы tbl_WEB_DiscountItem
                MySQLStr = "DELETE FROM tbl_WEB_DiscountItem "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---ну и собственно импорт новых значений
                StrCnt = 5
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '-----код группы товаров
                    MyCode = Trim(oSheet.getCellRangeByName("A" & StrCnt).String)
                    '---проверяем - есть ли такой товар в БД
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такой товар в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Товар с таким кодом отсутствует в базе. Импорт скидок по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            trycloseMyRec()
                        End If
                    End If
                    '-----Скидка
                    Try
                        MyDiscount = oSheet.getCellRangeByName("C" & StrCnt).Value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть числом" & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyDiscount <= 0 Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Скидка должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----занесение информации
                    Try
                        MySQLStr = "INSERT INTO tbl_WEB_DiscountItem "
                        MySQLStr = MySQLStr & "(ID, ItemCode, ClientCode, Discount) "
                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                        MySQLStr = MySQLStr & Replace(CStr(MyDiscount), ",", ".") & ")"

                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadAgreedRangeToExcel(ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка записей о согласованном ассортименте в Excel
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = MyCounter
        ExportAgreedRangeHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i, 0)
        ExportAgreedRangeBodyToExcel(MyWRKBook, MyCustomerCode, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadAgreedRangeToLO(ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByVal MyCounter As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка записей о согласованном ассортименте в LibreOffice
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = MyCounter
        '---Согласованный ассортимент
        ExportAgreedRangeHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i, 0)
        ExportAgreedRangeBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportAgreedRangeHeaderToExcel(ByRef MyWRKBook As Object, _
        ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByRef i As Integer, _
        ByVal WideFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка списка согласованного ассортимента в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Список согласованного ассортимента для клиента " & MyCustomerCodeName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Bold = True

        MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = MyCustomerCode
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font.Bold = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3)) = "Код товара Scala"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 3)) = "Название товара"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 3)) = "Цена (без НДС)"
        MyWRKBook.ActiveSheet.Range("D" & CStr(i + 3)) = "Код валюты"
        MyWRKBook.ActiveSheet.Range("E" & CStr(i + 3)) = "Валюта"

        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ":E" & CStr(i + 3)).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i + 3) & ";C" & CStr(i + 3) & ";D" & CStr(i + 3)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.599963377788629
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 15
        If WideFlag = 0 Then
            MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 15
        Else
            MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 80
        End If
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15

        i = i + 4
    End Function

    Public Function ExportAgreedRangeHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, _
        ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByRef i As Integer, _
        ByVal WideFlag As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка списка согласованного ассортимента в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2750
        If WideFlag = 0 Then
            oSheet.getColumns().getByName("D").Width = 2750
        Else
            oSheet.getColumns().getByName("D").Width = 15200
        End If
        oSheet.getColumns().getByName("E").Width = 2750

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список согласованного ассортимента для клиента " & MyCustomerCodeName
        oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), 10)

        oSheet.getCellRangeByName("C" & CStr(i)).String = MyCustomerCode
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))

        i = i + 3
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код группы товаров"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "имя группы товаров"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Цена (без НДС)"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Код валюты"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Валюта"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":A" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("B" & CStr(i) & ":B" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("D" & CStr(i) & ":D" & CStr(i)).CellBackColor = RGB(102, 255, 102)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("E" & CStr(i) & ":E" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))

        i = i + 1
    End Function

    Public Function ExportAgreedRangeBodyToExcel(ByRef MyWRKBook As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела согласованного ассортимента в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_AgreedRange.ItemCode, ISNULL(tbl_WEB_Items.Name, N'') AS Name, tbl_WEB_AgreedRange.AgreedPrice, "
        MySQLStr = MySQLStr & "tbl_WEB_AgreedRange.CurrCode, ISNULL(SYCD0100.SYCD009, N'') AS CurrName "
        MySQLStr = MySQLStr & "FROM tbl_WEB_AgreedRange INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON tbl_WEB_AgreedRange.CurrCode = SYCD0100.SYCD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_AgreedRange.ItemCode = tbl_WEB_Items.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_AgreedRange.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_AgreedRange.ItemCode "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportAgreedRangeBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByVal MyCustomerCode As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела согласованного ассортимента в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_WEB_AgreedRange.ItemCode, ISNULL(tbl_WEB_Items.Name, N'') AS Name, tbl_WEB_AgreedRange.AgreedPrice, "
        MySQLStr = MySQLStr & "tbl_WEB_AgreedRange.CurrCode, ISNULL(SYCD0100.SYCD009, N'') AS CurrName "
        MySQLStr = MySQLStr & "FROM tbl_WEB_AgreedRange INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON tbl_WEB_AgreedRange.CurrCode = SYCD0100.SYCD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_AgreedRange.ItemCode = tbl_WEB_Items.Code "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_AgreedRange.ClientCode = N'" & Trim(MyCustomerCode) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_AgreedRange.ItemCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(4)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CDbl(Declarations.MyRec.Fields(2).Value)
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub LoadAgreedRangeFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по согласованному ассортименту из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyPrice As Double
        Dim MyCurrCode As Integer
        Dim StrCnt As Double
        Dim MySQLStr As String

        MyTxt = "Для импорта данных по согласованному ассортименту вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружается согласованный ассортимент." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения цены (БЕЗ НДС) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должны располагаться коды валюты (Scala) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по согласованному ассортименту " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                '---------------Чтение и проверка кода клиента
                If appXLSRC.Worksheets(1).Range("C1").Value = Nothing Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    MyClientCode = Trim(appXLSRC.Worksheets(1).Range("C1").Value)
                    '---проверяем - есть ли такой клиент в БД, работающий через WEB
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                    MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            '---Очищаем старые значения из таблицы tbl_WEB_DiscountItem
                            MySQLStr = "DELETE FROM tbl_WEB_AgreedRange "
                            MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)

                            '---ну и собственно импорт новых значений
                            StrCnt = 5
                            While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                                Try
                                    MyCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                                    '---проверяем - есть ли такой товар в БД
                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                    MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
                                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                        trycloseMyRec()
                                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такой товар в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                        trycloseMyRec()
                                    Else
                                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Товар с таким кодом отсутствует в базе. Импорт скидок по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                            trycloseMyRec()
                                        Else
                                            trycloseMyRec()
                                            Try
                                                If appXLSRC.Worksheets(1).Range("C" & StrCnt).Value = Nothing Then
                                                    MyPrice = 0
                                                Else
                                                    MyPrice = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                                                End If
                                                If MyPrice <= 0 Then
                                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Согласованная цена без НДС должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                                                Else
                                                    Try
                                                        MyCurrCode = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                                                        '---проверяем - есть ли такой код валюты в БД
                                                        MySQLStr = "SELECT COUNT(*) AS CC "
                                                        MySQLStr = MySQLStr & "FROM SYCD0100 "
                                                        MySQLStr = MySQLStr & "WHERE (SYCD001 = " & MyCurrCode & ") "
                                                        MySQLStr = MySQLStr & "AND (SYCD001 IN (" & My.Settings.UsedCurr & ")) "
                                                        InitMyConn(False)
                                                        InitMyRec(False, MySQLStr)
                                                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                                            trycloseMyRec()
                                                            MsgBox("Ошибка в ячейке D" & StrCnt & ". Невозможно проверить, есть ли такой код валюты в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                                            trycloseMyRec()
                                                        Else
                                                            If Declarations.MyRec.Fields("CC").Value = 0 Then
                                                                MsgBox("Ошибка в ячейке D" & StrCnt & ". Такой код валюты отсутствует в базе или конфигурационном файле. Импорт согласованного ассортимента по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                                                                trycloseMyRec()
                                                            Else
                                                                trycloseMyRec()
                                                                Try
                                                                    MySQLStr = "INSERT INTO tbl_WEB_AgreedRange "
                                                                    MySQLStr = MySQLStr & "(ID, ItemCode, ClientCode, AgreedPrice, CurrCode) "
                                                                    MySQLStr = MySQLStr & "VALUES (NEWID(), "
                                                                    MySQLStr = MySQLStr & "N'" & MyCode & "', "
                                                                    MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                                                                    MySQLStr = MySQLStr & Replace(CStr(MyPrice), ",", ".") & ", "
                                                                    MySQLStr = MySQLStr & CStr(MyCurrCode) & ")"

                                                                    InitMyConn(False)
                                                                    Declarations.MyConn.Execute(MySQLStr)
                                                                Catch ex As Exception
                                                                    MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                                End Try
                                                            End If
                                                        End If
                                                    Catch ex As Exception
                                                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""D"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                                    End Try
                                                End If
                                            Catch ex As Exception
                                                MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                End Try
                                StrCnt = StrCnt + 1
                            End While
                        End If
                    End If
                End If

                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub LoadAgreedRangeFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по согласованному ассортименту из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MySQLStr As String
        Dim StrCnt As Double
        Dim MyClientCode As String
        Dim MyCode As String
        Dim MyPrice As Double
        Dim MyCurrCode As Integer

        MyTxt = "Для импорта данных по согласованному ассортименту вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке'C1' должен быть указан код клиента, для которого загружается согласованный ассортимент." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов (Скальские) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должны располагаться значения цены (БЕЗ НДС) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должны располагаться коды валюты (Scala) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Обратите внимание, что все старые данные по согласованному ассортименту " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "для данного клиента будут удалены и занесены только те, что есть в Excel файле." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '---------------Чтение и проверка кода клиента
                MyClientCode = oSheet.getCellRangeByName("C1").String
                If MyClientCode.Equals("") Then
                    MsgBox("Ошибка в ячейке ""C1"" - пустое значение. В ячейку ""С1"" необходимо занести код Скала покупателя вместе с предшествующими нолями.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                End If
                '---проверяем - есть ли такой клиент в БД, работающий через WEB
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & MyClientCode & "') "
                MySQLStr = MySQLStr & "AND (WorkOverWEB = 1) "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("Ошибка в ячейке ""C1"". Невозможно проверить, есть ли такой клиент в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    oWorkBook.Close(True)
                    Exit Sub
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке ""C1"". Клиент с таким кодом отсутствует в базе или данному клиенту не выставлен признак, что он работает через WEB. Импорт скидок по такому клиенту невозможен.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If
                End If

                '---Очищаем старые значения из таблицы tbl_WEB_DiscountItem
                MySQLStr = "DELETE FROM tbl_WEB_AgreedRange "
                MySQLStr = MySQLStr & "WHERE (ClientCode = N'" & MyClientCode & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---ну и собственно импорт новых значений
                StrCnt = 5
                While oSheet.getCellRangeByName("A" & StrCnt).String.Equals("") = False
                    '-----код товара
                    MyCode = Trim(oSheet.getCellRangeByName("A" & StrCnt).String)
                    '---проверяем - есть ли такой товар в БД
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
                    MySQLStr = MySQLStr & "WHERE (Code = N'" & MyCode & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке A" & StrCnt & ". Невозможно проверить, есть ли такой товар в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке A" & StrCnt & ". Товар с таким кодом отсутствует в базе. Импорт скидок по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            trycloseMyRec()
                        End If
                    End If
                    '-----Цена товара
                    Try
                        MyPrice = oSheet.getCellRangeByName("C" & StrCnt).Value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Цена должна быть числом" & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    If MyPrice <= 0 Then
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""C"". Цена должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '-----код валюты
                    Try
                        MyCurrCode = oSheet.getCellRangeByName("D" & StrCnt).Value
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""D"". Код валюты должен быть целым числом" & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    '---проверяем - есть ли такой код валюты в БД
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SYCD0100 "
                    MySQLStr = MySQLStr & "WHERE (SYCD001 = " & MyCurrCode & ") "
                    MySQLStr = MySQLStr & "AND (SYCD001 IN (" & My.Settings.UsedCurr & ")) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        trycloseMyRec()
                        MsgBox("Ошибка в ячейке D" & StrCnt & ". Невозможно проверить, есть ли такой код валюты в БД. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            MsgBox("Ошибка в ячейке D" & StrCnt & ". Такой код валюты отсутствует в базе или конфигурационном файле. Импорт согласованного ассортимента по такому товару невозможен.", MsgBoxStyle.Critical, "Внимание!")
                            trycloseMyRec()
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If

                    '-----занесение информации
                    Try
                        MySQLStr = "INSERT INTO tbl_WEB_AgreedRange "
                        MySQLStr = MySQLStr & "(ID, ItemCode, ClientCode, AgreedPrice, CurrCode) "
                        MySQLStr = MySQLStr & "VALUES (NEWID(), "
                        MySQLStr = MySQLStr & "N'" & MyCode & "', "
                        MySQLStr = MySQLStr & "N'" & MyClientCode & "', "
                        MySQLStr = MySQLStr & Replace(CStr(MyPrice), ",", ".") & ", "
                        MySQLStr = MySQLStr & CStr(MyCurrCode) & ")"

                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                        MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    StrCnt = StrCnt + 1
                End While
                oWorkBook.Close(True)
                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Public Sub UploadFULLDiscountsAgreedRangeToExcel(ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка записей о всех скидках и согласованном ассортименте в Excel
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 1
        ExportFULLDiscountsAgreedHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)

        '---скидки на группы
        ExportGroupDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportGroupDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)
        i = i + 2

        '---скидки на подгруппы
        ExportSubgroupDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportSubgroupDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)
        i = i + 2

        '---скидки на товары
        ExportItemDiscountHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportItemDiscountBodyToExcel(MyWRKBook, MyCustomerCode, i)
        i = i + 2

        '---согласованный ассортимент
        ExportAgreedRangeHeaderToExcel(MyWRKBook, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i, 1)
        ExportAgreedRangeBodyToExcel(MyWRKBook, MyCustomerCode, i)


        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadFULLDiscountsAgreedRangeToLO(ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка записей о всех скидках и согласованном ассортименте в LibreOffice
        '// MyCounter - счетчик строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportFULLDiscountsAgreedHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)

        '---скидки на группы
        ExportGroupDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportGroupDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)
        i = i + 2

        '---скидки на подгруппы
        ExportSubgroupDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportSubgroupDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)
        i = i + 2

        '---скидки на товары
        ExportItemDiscountHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i)
        ExportItemDiscountBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)
        i = i + 2

        '---согласованный ассортимент
        ExportAgreedRangeHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, MyCustomerCodeName, MyDiscount, MyPriceInfo, i, 1)
        ExportAgreedRangeBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, MyCustomerCode, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportFULLDiscountsAgreedHeaderToExcel(ByRef MyWRKBook As Object, _
        ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка записей о всех скидках и согласованном ассортименте в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Список скидок и согласованного ассортимента для клиента " & MyCustomerCodeName
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo

        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font.Bold = True

        With MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 1)).Font
            .ThemeColor = 10
            .TintAndShade = -0.249977111117893
        End With


        i = i + 3
    End Function

    Public Function ExportFULLDiscountsAgreedHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, _
        ByVal MyCustomerCode As String, _
        ByVal MyCustomerCodeName As String, _
        ByVal MyDiscount As Double, _
        ByVal MyPriceInfo As String, _
        ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка записей о всех скидках и согласованном ассортименте в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 1900
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2750

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Список скидок и согласованного ассортимента для клиента " & MyCustomerCodeName
        oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Общая скидка клиента " & CStr(MyDiscount) & " %, Прайс - лист: " & MyPriceInfo
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 1), 12)
        oSheet.getCellRangeByName("B" & CStr(i) & ":B" & CStr(i + 1)).CharColor = RGB(20, 20, 180) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        i = i + 3
    End Function

    Public Sub UploadBasePriceToExcel(ByVal MyPrice As String, ByVal MyPriceName As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка базового прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportBasePriceHeaderToExcel(MyWRKBook, i, MyPrice, MyPriceName)
        ExportBasePriceBodyToExcel(MyWRKBook, i, MyPrice, SubgroupFlag)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadBasePriceToLO(ByVal MyPrice As String, ByVal MyPriceName As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка базового прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportBasePriceHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyPrice, MyPriceName)
        ExportBasePriceBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyPrice, SubgroupFlag)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportBasePriceHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyPrice As String, ByVal MyPriceName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка базового прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Базовый прайс лист " & MyPriceName
        MyWRKBook.ActiveSheet.Range("B2") = "Для работы через WEB сайт на дату " & Format(Now(), "dd/MM/yyyy")

        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A4") = "Код товара"
        MyWRKBook.ActiveSheet.Range("B4") = "Название товара"
        MyWRKBook.ActiveSheet.Range("C4") = "Код производителя"
        MyWRKBook.ActiveSheet.Range("D4") = "Производитель"
        MyWRKBook.ActiveSheet.Range("E4") = "Код товара производителя"
        MyWRKBook.ActiveSheet.Range("F4") = "Цена (без НДС)"
        MyWRKBook.ActiveSheet.Range("G4") = "Валюта"

        MyWRKBook.ActiveSheet.Range("A4:G4").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A4:G4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A4:G4").WrapText = True
        MyWRKBook.ActiveSheet.Range("A4:G4").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:G4").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:G4").Select()
        MyWRKBook.ActiveSheet.Range("A4:G4").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A4:G4").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A4:G4").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:G4").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:G4").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:G4").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:G4").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:G4").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 11
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 11
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10

        i = 5
    End Function

    Public Function ExportBasePriceHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyPrice As String, ByVal MyPriceName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка базового прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2750
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2090
        oSheet.getColumns().getByName("D").Width = 5700
        oSheet.getColumns().getByName("E").Width = 2750
        oSheet.getColumns().getByName("F").Width = 2750
        oSheet.getColumns().getByName("G").Width = 1900

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Базовый прайс лист " & MyPriceName
        i = 2
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Для работы через WEB сайт на дату " & Format(Now(), "dd/MM/yyyy")
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), 10)

        i = 4
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код товара"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Название товара"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Код производителя"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Производитель"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Код товара производителя"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Цена (без НДС)"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Валюта"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":G" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":G" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":G" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":G" & CStr(i))

        i = 5
    End Function

    Public Function ExportBasePriceBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyPrice As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела базового прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, '') AS ManufacturerName, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(tbl_WEB_Items.ManufacturerItemCode)) AS ManufacturerItemCode, SC390300.SC39005, SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
        MySQLStr = MySQLStr & "SC390300 ON tbl_WEB_Items.Code = SC390300.SC39001 INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON SC390300.SC39003 = SYCD0100.SYCD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
        MySQLStr = MySQLStr & "WHERE (SC390300.SC39002 = N'" & Trim(MyPrice) & "') "
        If SubgroupFlag = True Then
            MySQLStr = MySQLStr & "AND (Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) <> N'') "
        End If
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportBasePriceBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyPrice As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела базового прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, '') AS ManufacturerName, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(tbl_WEB_Items.ManufacturerItemCode)) AS ManufacturerItemCode, SC390300.SC39005, SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
        MySQLStr = MySQLStr & "SC390300 ON tbl_WEB_Items.Code = SC390300.SC39001 INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON SC390300.SC39003 = SYCD0100.SYCD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
        MySQLStr = MySQLStr & "WHERE (SC390300.SC39002 = N'" & Trim(MyPrice) & "') "
        If SubgroupFlag = True Then
            MySQLStr = MySQLStr & "AND (Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) <> N'') "
        End If
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(6)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CInt(Declarations.MyRec.Fields(2).Value)
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = CDbl(Declarations.MyRec.Fields(5).Value)
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":G" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Sub UploadIndividualPriceToExcel(ByVal MyCustomer As String, ByVal MyCustomerName As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка индивидуального прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportIndividualPriceHeaderToExcel(MyWRKBook, i, MyCustomer, MyCustomerName)
        ExportindividualPriceBodyToExcel(MyWRKBook, i, MyCustomer, SubgroupFlag)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Public Sub UploadIndividualPriceToLO(ByVal MyCustomer As String, ByVal MyCustomerName As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка индивидуального прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportIndividualPriceHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyCustomer, MyCustomerName)
        ExportindividualPriceBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyCustomer, SubgroupFlag)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Public Function ExportIndividualPriceHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyCustomer As String, ByVal MyCustomerName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка индивидуального прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Индивидуальный прайс лист для " & MyCustomerName
        MyWRKBook.ActiveSheet.Range("B2") = "Для работы через WEB сайт на дату " & Format(Now(), "dd/MM/yyyy")

        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A4") = "Код товара"
        MyWRKBook.ActiveSheet.Range("B4") = "Название товара"
        MyWRKBook.ActiveSheet.Range("C4") = "Код производителя"
        MyWRKBook.ActiveSheet.Range("D4") = "Производитель"
        MyWRKBook.ActiveSheet.Range("E4") = "Код товара производителя"
        MyWRKBook.ActiveSheet.Range("F4") = "Вид скидки"
        MyWRKBook.ActiveSheet.Range("G4") = "Скидка"
        MyWRKBook.ActiveSheet.Range("H4") = "Себестоимость"
        MyWRKBook.ActiveSheet.Range("I4") = "Цена Руб (без НДС)"
        MyWRKBook.ActiveSheet.Range("J4") = "Маржа"

        MyWRKBook.ActiveSheet.Range("A4:J4").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A4:J4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A4:J4").WrapText = True
        MyWRKBook.ActiveSheet.Range("A4:J4").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:J4").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:J4").Select()
        MyWRKBook.ActiveSheet.Range("A4:J4").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A4:J4").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 11
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 10

        i = 5
    End Function

    Public Function ExportIndividualPriceHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyCustomer As String, ByVal MyCustomerName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка индивидуального прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2750
        oSheet.getColumns().getByName("B").Width = 15200
        oSheet.getColumns().getByName("C").Width = 2090
        oSheet.getColumns().getByName("D").Width = 5700
        oSheet.getColumns().getByName("E").Width = 2750
        oSheet.getColumns().getByName("F").Width = 3850
        oSheet.getColumns().getByName("G").Width = 1900
        oSheet.getColumns().getByName("H").Width = 1950
        oSheet.getColumns().getByName("I").Width = 1950
        oSheet.getColumns().getByName("J").Width = 1900

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Индивидуальный прайс лист для " & MyCustomerName
        i = 2
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Для работы через WEB сайт на дату " & Format(Now(), "dd/MM/yyyy")
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), 10)

        i = 4
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код товара"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Название товара"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Код производителя"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Производитель"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Код товара производителя"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Вид скидки"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Скидка"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "Себестоимость"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Цена Руб (без НДС)"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "Маржа"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))

        i = 5
    End Function

    Public Function ExportindividualPriceBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyCustomer As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела индивидуального прайс листа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_WEB_IndividualPricePreparation N'" & Trim(MyCustomer) & "', " & IIf(SubgroupFlag = True, 1, 0)

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportindividualPriceBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyCustomer As String, ByVal SubgroupFlag As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела индивидуального прайс листа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "exec spp_WEB_IndividualPricePreparation N'" & Trim(MyCustomer) & "', " & IIf(SubgroupFlag = True, 1, 0)

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(9)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CInt(Declarations.MyRec.Fields(2).Value)
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = CDbl(Declarations.MyRec.Fields(6).Value)
                MyArr(7) = CDbl(Declarations.MyRec.Fields(7).Value)
                MyArr(8) = CDbl(Declarations.MyRec.Fields(8).Value)
                MyArr(9) = CDbl(Declarations.MyRec.Fields(9).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Function UploadItemDimToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel
        '// с информацией по длине, ширине и высоте
        '// а также по весу
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 0
        ExportItemDimHeaderToExcel(MyWRKBook, i)
        ExportItemDimBodyToExcel(MyWRKBook, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Function

    Public Function UploadItemDimToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в LibreOffice
        '// с информацией по длине, ширине и высоте
        '// а также по весу
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportItemDimHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        ExportItemDimBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Function

    Public Function ExportItemDimHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка информации по длине, ширине, высоте и весу товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "Информация по измерениям и весу товаров "
        MyWRKBook.ActiveSheet.Range("B2") = "На дату " & Format(Now(), "dd/MM/yyyy")

        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("A4") = "Код товара"
        MyWRKBook.ActiveSheet.Range("B4") = "Длина"
        MyWRKBook.ActiveSheet.Range("C4") = "Ширина"
        MyWRKBook.ActiveSheet.Range("D4") = "Высота"
        MyWRKBook.ActiveSheet.Range("E4") = "Вес"
        MyWRKBook.ActiveSheet.Range("F4") = "Код поставщика"
        MyWRKBook.ActiveSheet.Range("G4") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("H4") = "Код производителя"
        MyWRKBook.ActiveSheet.Range("I4") = "Название производителя"
        MyWRKBook.ActiveSheet.Range("J4") = "Код товара производителя"

        MyWRKBook.ActiveSheet.Range("A4:J4").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A4:J4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A4:J4").WrapText = True
        MyWRKBook.ActiveSheet.Range("A4:J4").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:J4").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A4:J4").Select()
        MyWRKBook.ActiveSheet.Range("A4:J4").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A4:J4").Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:J4").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 20

        i = 5
    End Function

    Public Function ExportItemDimHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка информации по длине, ширине, высоте и весу товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2750
        oSheet.getColumns().getByName("B").Width = 3800
        oSheet.getColumns().getByName("C").Width = 3800
        oSheet.getColumns().getByName("D").Width = 3800
        oSheet.getColumns().getByName("E").Width = 3800
        oSheet.getColumns().getByName("F").Width = 3800
        oSheet.getColumns().getByName("G").Width = 3800
        oSheet.getColumns().getByName("H").Width = 3800
        oSheet.getColumns().getByName("I").Width = 3800
        oSheet.getColumns().getByName("J").Width = 3800

        oSheet.getCellRangeByName("B" & CStr(i)).String = "Информация по измерениям и весу товаров"
        i = 2
        oSheet.getCellRangeByName("B" & CStr(i)).String = "На дату " & Format(Now(), "dd/MM/yyyy")
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i - 1) & ":B" & CStr(i), 10)

        i = 4
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код товара"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Длина"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Ширина"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Высота"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Вес"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Код поставщика"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Код товара поставщика"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "Код производителя"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Название производителя"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "Код товара производителя"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))

        i = 5
    End Function

    Public Function ExportItemDimBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела информации по длине, ширине, высоте и весу товаров в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT SC010300.SC01001, SC010300.SC01007, SC010300.SC01008, SC010300.SC01009, SC010300.SC01069, SC010300.SC01058, SC010300.SC01060, "
        MySQLStr = MySQLStr & "tbl_ItemCard0300.Manufacturer, tbl_Manufacturers.Name, tbl_ItemCard0300.ManufacturerItemCode "
        MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Public Function ExportItemDimBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела информации по длине, ширине, высоте и весу товаров в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT SC010300.SC01001, SC010300.SC01007, SC010300.SC01008, SC010300.SC01009, SC010300.SC01069, SC010300.SC01058, SC010300.SC01060, "
        MySQLStr = MySQLStr & "tbl_ItemCard0300.Manufacturer, tbl_Manufacturers.Name, tbl_ItemCard0300.ManufacturerItemCode "
        MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(9)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = CDbl(Declarations.MyRec.Fields(1).Value)
                MyArr(2) = CDbl(Declarations.MyRec.Fields(2).Value)
                MyArr(3) = CDbl(Declarations.MyRec.Fields(3).Value)
                MyArr(4) = CDbl(Declarations.MyRec.Fields(4).Value)
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArr(7) = CInt(Declarations.MyRec.Fields(7).Value)
                MyArr(8) = Declarations.MyRec.Fields(8).Value
                MyArr(9) = Declarations.MyRec.Fields(9).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If
    End Function

    Public Function LoadItemDimFromExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продуктам (длина, ширина, высота и вес) из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim appXLSRC As Object
        Dim MyScalaCode As String
        Dim MyLength As Double
        Dim MyWidth As Double
        Dim MyHeight As Double
        Dim MyWeight As Double
        Dim MySQLStr As String
        Dim StrCnt As String

        MyTxt = "Для импорта данных по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды товара в Scala) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов Scala (с предшествующими нулями, если есть) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'B' должна быть размещена длина товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должна быть размещена ширина товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должна быть размещена высота товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' должен быть размещен вес товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog1.ShowDialog()
            If (MainForm.OpenFileDialog1.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(MainForm.OpenFileDialog1.FileName)

                StrCnt = 5
                While Not appXLSRC.Worksheets(1).Range("A" & StrCnt).Value = Nothing
                    MyScalaCode = Trim(appXLSRC.Worksheets(1).Range("A" & StrCnt).Value)
                    If Trim(MyScalaCode) <> "" Then
                        Try
                            MyLength = appXLSRC.Worksheets(1).Range("B" & StrCnt).Value
                            MyWidth = appXLSRC.Worksheets(1).Range("C" & StrCnt).Value
                            MyHeight = appXLSRC.Worksheets(1).Range("D" & StrCnt).Value
                            MyWeight = appXLSRC.Worksheets(1).Range("E" & StrCnt).Value

                            '---Запись нового значения
                            MySQLStr = "UPDATE SC010300 "
                            MySQLStr = MySQLStr & "SET SC01007 = " & Replace(CStr(MyLength), ",", ".") & ", "
                            MySQLStr = MySQLStr & "SC01008 = " & Replace(CStr(MyWidth), ",", ".") & ", "
                            MySQLStr = MySQLStr & "SC01009 = " & Replace(CStr(MyHeight), ",", ".") & ", "
                            MySQLStr = MySQLStr & "SC01069 = " & Replace(CStr(MyWeight), ",", ".") & " "
                            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Replace(MyScalaCode, "'", "''") & "')"
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)

                        Catch ex As Exception
                            MsgBox("Ошибка в строке " & StrCnt & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        End Try

                    Else
                        MsgBox("Ошибка в строке " & StrCnt & " колонке ""A"". Указание кода товара Scala обязательно.", MsgBoxStyle.Critical, "Внимание!")
                    End If
                    StrCnt = StrCnt + 1
                End While
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing

                MainForm.Cursor = Cursors.Default
            End If
        End If
    End Function

    Public Function LoadItemDimFromLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка информации по продуктам (длина, ширина, высота и вес) из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MySQLStr As String
        Dim MyTableName As String                   'Имя временной таблицы
        Dim MyGuid As String                          '
        Dim oCol As Object              '---Колонка, в которой смотрим диапазоны
        Dim oBlank As Object            '---набор пустых диапазонов
        Dim oRg                         '---пустой диапазон
        Dim oRange As Object
        Dim EndRange As Integer         '---Конец заполненного диапазона (строка начала пустого диапазона (ID))
        Dim StartRange As Integer
        Dim MyArr() As Object
        Dim MySQLAdapter As SqlClient.SqlDataAdapter 'для временной таблицы
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyDbl As Double

        MyTxt = "Для импорта данных по продуктам вам необходимо подготовить файл Excel. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Данные в файле Excel должны начинаться с 5 строки, с колонки 'A'." & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки в колонке 'A' (коды товара в Scala) должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'A' должны располагаться коды продуктов Scala (с предшествующими нулями, если есть) " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'B' должна быть размещена длина товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'C' должна быть размещена ширина товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'D' должна быть размещена высота товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке 'E' должен быть размещен вес товара " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
        If (MyRez = MsgBoxResult.Ok) Then
            MainForm.OpenFileDialog2.ShowDialog()
            If (MainForm.OpenFileDialog2.FileName = "") Then
            Else
                MainForm.Cursor = Cursors.WaitCursor
                System.Windows.Forms.Application.DoEvents()

                oServiceManager = CreateObject("com.sun.star.ServiceManager")
                oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                oFileName = Replace(MainForm.OpenFileDialog2.FileName, "\", "/")
                oFileName = "file:///" + oFileName
                Dim arg(1)
                arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                oSheet = oWorkBook.getSheets().getByIndex(0)

                '-----Вариант с импортом во временную таблицу
                '---определение диапазона данных
                StartRange = 5
                oCol = oSheet.Columns.getByIndex(0)
                oBlank = oCol.queryEmptyCells()
                oRg = oBlank.getByIndex(1)
                EndRange = oRg.RangeAddress.StartRow

                MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
                MyTableName = "tbl_ItemsDimension_Tmp_" + MyGuid
                '---Создание временных таблиц
                Try
                    MySQLStr = "DROP TABLE " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex As Exception
                End Try
                MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35) NOT NULL, "
                MySQLStr = MySQLStr & "[ItemLength] [numeric](18, 8) NOT NULL, "
                MySQLStr = MySQLStr & "[ItemWidth] [numeric](18, 8) NOT NULL, "
                MySQLStr = MySQLStr & "[ItemHeight] [numeric](18, 8) NOT NULL, "
                MySQLStr = MySQLStr & "[ItemWeight] [numeric](18, 8) NOT NULL "
                MySQLStr = MySQLStr & "CONSTRAINT [PK_" & MyTableName & "] PRIMARY KEY CLUSTERED "
                MySQLStr = MySQLStr & "( "
                MySQLStr = MySQLStr & "[ItemCode] ASC "
                MySQLStr = MySQLStr & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, "
                MySQLStr = MySQLStr & "ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY] "
                MySQLStr = MySQLStr & ") ON [PRIMARY] "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '-----данные 
                InitMyConn(False)
                MySQLStr = "SELECT ItemCode, ItemLength, ItemWidth, ItemHeight, ItemWeight "
                MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                Try
                    MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MySQLAdapter.SelectCommand.CommandTimeout = 1200
                    Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                    MySQLAdapter.Fill(MySQLDs)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex1 As Exception
                    End Try
                    oWorkBook.Close(True)
                    Exit Function
                End Try

                '-----Перенос данных из Excel dataset в SQL dataset
                Dim dt As DataTable
                Dim dr As DataRow

                dt = MySQLDs.Tables(0)
                oRange = oSheet.getCellRangeByName("A" & CStr(StartRange) & ":E" & CStr(EndRange))
                MyArr = oRange.DataArray
                For i As Integer = 0 To EndRange - 6
                    dr = dt.NewRow
                    '---код скала
                    If MyArr(i)(0).Equals("") Then
                        Exit For
                    End If
                    dr.Item(0) = MyArr(i)(0)
                    '---длина
                    Try
                        MyDbl = MyArr(i)(1)
                        dr.Item(1) = MyArr(i)(1)
                    Catch ex As Exception
                        MsgBox("Ячейка B" & CStr(i) & " должно быть числовое значение")
                        Try
                            MySQLStr = "DROP TABLE " & MyTableName & " "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        Catch ex1 As Exception
                        End Try
                        oWorkBook.Close(True)
                        Exit Function
                    End Try
                    '---ширина
                    Try
                        MyDbl = MyArr(i)(2)
                        dr.Item(2) = MyArr(i)(2)
                    Catch ex As Exception
                        MsgBox("Ячейка C" & CStr(i) & " должно быть числовое значение")
                        Try
                            MySQLStr = "DROP TABLE " & MyTableName & " "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        Catch ex1 As Exception
                        End Try
                        oWorkBook.Close(True)
                        Exit Function
                    End Try
                    '---высота
                    Try
                        MyDbl = MyArr(i)(3)
                        dr.Item(3) = MyArr(i)(3)
                    Catch ex As Exception
                        MsgBox("Ячейка D" & CStr(i) & " должно быть числовое значение")
                        Try
                            MySQLStr = "DROP TABLE " & MyTableName & " "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        Catch ex1 As Exception
                        End Try
                        oWorkBook.Close(True)
                        Exit Function
                    End Try
                    '---вес
                    Try
                        MyDbl = MyArr(i)(4)
                        dr.Item(4) = MyArr(i)(4)
                    Catch ex As Exception
                        MsgBox("Ячейка E" & CStr(i) & " должно быть числовое значение")
                        Try
                            MySQLStr = "DROP TABLE " & MyTableName & " "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        Catch ex1 As Exception
                        End Try
                        oWorkBook.Close(True)
                        Exit Function
                    End Try
                    dt.Rows.Add(dr)
                Next i
                Try
                    MySQLAdapter.Update(MySQLDs, "Table")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex1 As Exception
                    End Try
                    oWorkBook.Close(True)
                    Exit Function
                End Try

                '-----Обновление данных в таблице SC010300
                MySQLStr = "UPDATE SC010300 "
                MySQLStr = MySQLStr & "SET SC01007 = " & MyTableName & ".ItemLength, "
                MySQLStr = MySQLStr & "SC01008 = " & MyTableName & ".ItemWidth, "
                MySQLStr = MySQLStr & "SC01009 = " & MyTableName & ".ItemHeight, "
                MySQLStr = MySQLStr & "SC01069 = " & MyTableName & ".ItemWeight "
                MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
                MySQLStr = MySQLStr & MyTableName & " ON SC010300.SC01001 = " & MyTableName & ".ItemCode"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                Try
                    MySQLStr = "DROP TABLE " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex1 As Exception
                End Try
                oWorkBook.Close(True)
            End If
        End If
    End Function
End Module
