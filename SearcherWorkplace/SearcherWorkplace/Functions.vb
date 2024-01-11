Module Functions
    Public Sub InitMyConn(ByVal IsSystem As Boolean)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Инициализация соединения с БД, чтение глобальных переменных
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim Scala As New SfwIII.Application

        On Error GoTo MyCatch
        If MyConn Is Nothing Then
            MyConn = New ADODB.Connection
            MyConn.CursorLocation = 3
            MyConn.CommandTimeout = 600
            MyConn.ConnectionTimeout = 300
            If Declarations.MyConnStr = "" Then
                Declarations.MyConnStr = Scala.ActiveProcess.UserContext.GetConnectionString(1)
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SQLCLS"
                Declarations.MyNETConnStr = Replace(Declarations.MyConnStr, "Provider=SQLOLEDB;", "")
                Declarations.MyNETConnStr = Declarations.MyNETConnStr & ";Timeout=0;"
            End If
            If IsSystem = True Then
                MyConn.Open(Replace(Declarations.MyConnStr, "ScaDataDB", "ScalaSystemDB"))
            Else
                MyConn.Open(Declarations.MyConnStr)
            End If
            If Declarations.CompanyID = "" Then
                Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            End If
            If Declarations.Year = "" Then
                Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            End If
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 1")
    End Sub

    Public Sub trycloseMyRec()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Попытка закрытия рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        On Error Resume Next
        MyRec.Close()
    End Sub

    Public Sub InitMyRec(ByVal IsSystem As Boolean, ByVal sql As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Открытие рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyErr

        On Error GoTo MyCatch
        InitMyConn(IsSystem)
        If MyRec Is Nothing Then
            MyRec = New ADODB.Recordset
        End If
        trycloseMyRec()
        MyRec.LockType = LockTypeEnum.adLockOptimistic
        MyRec.Open(sql, MyConn)
        If MyConn.Errors.Count > 0 Then
            For Each MyErr In MyConn.Errors
                Err.Raise(MyErr.Number, MyErr.Source, MyErr.Description)
            Next MyErr
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 2")
    End Sub

    Public Function UpdateRequestDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// обновление строк предложения из Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim appXLSRC As Object
        Dim MyExcelCounter As Double                    'счетчик строк Excel
        Dim MyItemCode As String                        'код товара Scala
        Dim MySuppItemCode As String                    'код товара поставщика
        Dim MyItemName As String                        'название товара в запросе
        Dim MyItemCounter As Double                     'счетчик кол - ва Скальских кодов товара 
        Dim MyUOM As Integer                            'единица измерения
        Dim c As Object
        Dim MyQTY As Double
        Dim MyPrice As Double
        Dim MyWeekQTY As Double
        Dim MyCurr As Integer                           'код валюты

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            MyExcelCounter = 11
            MyCurr = 0

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing

                '------код запаса Scala
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    MyItemCode = ""
                Else
                    If Trim(Declarations.MySupplierCode).Equals("") Then
                        MsgBox("Ячейка B" & MyExcelCounter & ". Вы пытаетесь обновить код товара в Scala для поставщика " & Trim(Declarations.MySupplierName) & _
                            ", который в Scala не заведен. Сначала заведите в Scala поставщика или замените в списке поставщиков для поиска поставщика на поставщика из SCala.", _
                            MsgBoxStyle.Critical, "Внимание!")
                        Exit While
                    Else
                        MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                        '---проверка что такой код есть в Scala у выбранного поставщика
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                        MySQLStr = MySQLStr & "AND (SC01058 = N'" & Trim(Declarations.MySupplierCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("Ячейка B" & MyExcelCounter & ". Невозможно получить информацию о запасах из Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyItemCounter = Declarations.MyRec.Fields("CC").Value
                            trycloseMyRec()
                        End If
                        If MyItemCounter = 0 Then
                            MsgBox("Ячейка B" & MyExcelCounter & ". такого кода товара у поставщика " & Trim(Declarations.MySupplierName) & " в Scala нет.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        Else
                        End If
                    End If
                End If

                '------код товара поставщика
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                End If

                '------Название товара
                If appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing Then
                    MyItemName = ""
                Else
                    MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value.ToString
                End If

                '-------Проверка заполнения первых 3 - х полей
                If Trim(MyItemCode) = "" And (Trim(MySuppItemCode) = "" Or Trim(MyItemName) = "") Then
                    MsgBox("Строка " & MyExcelCounter & " Обязательно должен быть занесен или код товара Scala, или код товара поставщика + название товара!", MsgBoxStyle.Critical, "Внимание!")
                    Exit While
                Else

                    '------Единица измерения
                    If Trim(MyItemCode).Equals("") Then
                        c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                        If c.Text = "" Then
                            MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' не выбрано из списка формы. Необходимо выбрать значение из выпадающего списка.", MsgBoxStyle.YesNo, "Внимание!")
                            Exit While
                        Else
                            MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                        End If
                    Else
                        MySQLStr = "SELECT SC01135 AS UOM "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("Ячейка B" & MyExcelCounter & ". Невозможно получить информацию об единице измерения этого товара из Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyUOM = Declarations.MyRec.Fields("UOM").Value
                            trycloseMyRec()
                        End If
                    End If

                    '------Количество
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("Ячейка F" & MyExcelCounter & ". Количество должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
                        Exit While
                    Else
                        MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                        If MyQTY = 0 Then
                            MsgBox("Ячейка F" & MyExcelCounter & ". Количество должно быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        End If
                    End If
                    '------Закупочная цена без НДС
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("Ячейка G" & MyExcelCounter & ". Закупочная цена без НДС должна быть заполнена.", MsgBoxStyle.Critical, "Внимание!")
                        Exit While
                    Else
                        MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                        If MyPrice = 0 Then
                            MsgBox("Ячейка G" & MyExcelCounter & ". Закупочная цена без НДС должна быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        End If
                    End If
                    '----------Срок поставки
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("Ячейка I" & MyExcelCounter & ". Срок поставки должен быть заполнен.", MsgBoxStyle.Critical, "Внимание!")
                        Exit While
                    Else
                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                        If MyWeekQTY = 0 Then
                            MsgBox("Ячейка I" & MyExcelCounter & ". Срок поставки должен быть больше нуля.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        End If
                    End If

                    '-----Обновление информации
                    ''-----по коду товара производителя
                    'If MySuppItemCode.Equals("") = False Then
                    '    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    '    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    '    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    '    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    '    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    '    MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                    '    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    '    InitMyConn(False)
                    '    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                    ''-----по названию  товара
                    'If MyItemName.Equals("") = False Then
                    '    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    '    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    '    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    '    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    '    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    '    MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                    '    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    '    InitMyConn(False)
                    '    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                    '-----по коду товара производителя + названию  товара
                    'If MySuppItemCode.Equals("") = False Then
                    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                    MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                End If
                    MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------удаление объектов
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("Процедура обновления запроса завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Function UpdateRequestDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// обновление строк предложения из LibreOffice
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyItemCode As String                        'код товара Scala
        Dim MySuppItemCode As String                  'код товара поставщика
        Dim MyItemName As String                      'название товара
        Dim MyUOMStr As String                        'единица измерения
        Dim MyUOM As Integer                          'код единицы измерения
        Dim MyQTY As Double                           'кол - во
        Dim MyPrice As Double                         'цена за 1
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       'Версия документа
        Dim MySQLStr As String                        'SQL запрос
        Dim MyRez As Object
        Dim MyExcelCounter As Double                  'счетчик строк Excel
        Dim MyItemCounter As Double                   'счетчик кол - ва Скальских кодов товара соотв. коду товара поставщика
        Dim MyScalaItemName As String                 'название товара в Scala
        Dim MyScalaUOM As Integer                     'код единицы измерения в Scala
        Dim MyScalaSuppItemCode As String             'код товара поставщика в Scala
        Dim MyWeekQTY As Double                       'срок поставки

        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        oFileName = Replace(Declarations.ImportFileName, "\", "/")
        oFileName = "file:///" + oFileName
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        MyExcelCounter = 11
        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------код запаса Scala
            MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            If MyItemCode.Equals("") = False Then
                '---проверка что такой код есть в Scala
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MyItemCounter = 0
                Else
                    Declarations.MyRec.MoveFirst()
                    MyItemCounter = Declarations.MyRec.Fields("CC").Value
                    trycloseMyRec()
                End If
                If MyItemCounter = 0 Then
                    MyItemCode = ""
                Else
                    '-----получаем информацию по скальскому товару
                    MySQLStr = "SELECT SC01135 AS UOM, LTRIM(RTRIM(LTRIM(RTRIM(SC01002)) + ' ' + LTRIM(RTRIM(SC01003)))) AS Name, SC01060 AS SuppItemID "
                    MySQLStr = MySQLStr & "FROM SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "

                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)

                    MyScalaUOM = CInt(Declarations.MyRec.Fields("UOM").Value)
                    MyScalaItemName = Declarations.MyRec.Fields("Name").Value
                    MyScalaSuppItemCode = Declarations.MyRec.Fields("SuppItemID").Value
                    trycloseMyRec()
                End If
            End If
            '------код запаса поставщика
            MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
            '------Название товара
            MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
            '-------Проверка заполнения первых 3 - х полей
            '-----код Scala не заносим
            If Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                MsgBox("Строка " & MyExcelCounter & " Обязательно должен быть занесен код товара поставщика или название товара!", MsgBoxStyle.Critical, "Внимание!")
                Exit While
            Else
                '------единица измерения
                MyUOMStr = oSheet.getCellRangeByName("E" & MyExcelCounter).String
                Dim MyRange As Object = oSheet.getCellrangeByName("O11:O18")
                Dim Search_Desc As Object = MyRange.createSearchDescriptor()
                Search_Desc.SearchString = MyUOMStr
                Dim Search_Result As Object = MyRange.findAll(Search_Desc)
                If Search_Result.Count < 1 Then
                    MyRez = MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' не выбрано из списка формы. Необходимо выбрать значение из выпадающего списка. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    Dim Last_Occur As Object = Search_Result.getByIndex(Search_Result.Count - 1)
                    MyUOM = oSheet.getCellRangeByName("N" & Last_Occur.CellAddress.Row + 1).String
                End If
                '-----количество
                Try
                    MyQTY = oSheet.getCellRangeByName("F" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("Ячейка F" & MyExcelCounter & " значение 'Количество' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '----------Цена
                Try
                    MyPrice = oSheet.getCellRangeByName("G" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("Ячейка G" & MyExcelCounter & " значение 'Цена без НДС' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '----------Срок поставки
                Try
                    MyWeekQTY = oSheet.getCellRangeByName("I" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("Ячейка I" & MyExcelCounter & " значение 'Срок поставки' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '-----Обновление информации
                '-----по коду товара производителя + названию  товара
                'If MySuppItemCode.Equals("") = False Then
                MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                MySQLStr = MySQLStr & ",CurrCode = 0 "
                MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                'End If
            End If

            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("Процедура обновления запроса завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Function CheckVersion(ByVal MyVer As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка версии Excel файла - можно ли с ней работать
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            CheckVersion = False
            trycloseMyRec()
            Exit Function
        Else
            If Trim(Declarations.MyRec.Fields("Version").Value) = Trim(MyVer) Then
                CheckVersion = True
                trycloseMyRec()
                Exit Function
            Else
                MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Trim(Declarations.MyRec.Fields("Version").Value) & ".", vbCritical, "Внимание!")
                CheckVersion = False
                trycloseMyRec()
                Exit Function
            End If
        End If
    End Function

    Public Function GetEmailFromDB(ByVal UserCode As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение Email по коду продавца / закупщика...
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT RM66003 "
        MySQLStr = MySQLStr & "FROM RM.dbo.RM660100 "
        MySQLStr = MySQLStr & "WHERE (RM66001 = '" & Trim(UserCode) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            GetEmailFromDB = ""
        Else
            GetEmailFromDB = Trim(Declarations.MyRec.Fields("RM66003").Value.ToString)
        End If
        trycloseMyRec()
    End Function

    Public Sub SendInfoByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewRequestState As String, ByVal SearcherName As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка по почте уведомления продавцам об изменениях в запросах
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String
        Dim MailTo() As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        MailTo = Split(EMailTo, ";")
        For i As Integer = 0 To UBound(MailTo)
            msg.To.Add(MailTo(i))
        Next
        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "Уведомление об изменении запроса на поиск"
        MyMsgStr = "Уважаемый коллега!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "Произошло изменение следующего запроса на поиск:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " от " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "Для клиента: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "Продавец: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "Поисковик: " & SearcherName & Chr(13)
        MyMsgStr = MyMsgStr & "Новое состояние запроса: " & NewRequestState & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "С уважением," & Chr(13)
        MyMsgStr = MyMsgStr + "ООО ""Скандика"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. На письмо просьба не отвечать, это автоматическая рассылка. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

    Public Sub SendCommentByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewComment As String, ByVal SearcherName As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка по почте уведомления продавцам о новом комментарии
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String
        Dim MailTo() As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        MailTo = Split(EMailTo, ";")
        For i As Integer = 0 To UBound(MailTo)
            msg.To.Add(MailTo(i))
        Next
        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "Уведомление о новом комментарии к запросу на поиск"
        MyMsgStr = "Уважаемый коллега!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "Получен новый комментарий к следующему запросу на поиск:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " от " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "Для клиента: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "Продавец: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "Поисковик: " & SearcherName & Chr(13)
        MyMsgStr = MyMsgStr & "Новый комментарий к запросу: " & NewComment & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "С уважением," & Chr(13)
        MyMsgStr = MyMsgStr + "ООО ""Скандика"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. На письмо просьба не отвечать, это автоматическая рассылка. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub
End Module
