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
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SPBDVL3"
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=sqlcls"
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

    Public Function GetNewID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение следующего свободного ID предложения на продажу 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyID As Double                            'ID
        Dim MyTID As Double                           '

        Do
            MyID = GetNextID()
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            MyTID = Declarations.MyRec.Fields("CC").Value
        Loop While MyTID <> 0
        GetNewID = MyID
    End Function

    Public Function GetNextID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение следующего ID предложения на продажу 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyID As Double

        MySQLStr = "Select SY68002 "
        MySQLStr = MySQLStr & "FROM tbl_SY6803XX "
        MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyID = 1
            trycloseMyRec()
        Else
            MyID = Declarations.MyRec.Fields("SY68002").Value
            trycloseMyRec()
            MySQLStr = "UPDATE tbl_SY6803XX "
            MySQLStr = MySQLStr & "SET SY68002 = " & CStr(MyID + 1) & " "
            MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
            Declarations.MyConn.Execute(MySQLStr)
        End If
        GetNextID = MyID
    End Function

    Public Function GetNewPRDID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение следующего свободного ID заказа на продажу 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyID As Double                            'ID
        Dim MyTID As Double                           '

        Do
            MyID = GetNextPRDID()
            'MySQLStr = "SELECT COUNT(*) AS CC "
            'MySQLStr = MySQLStr & "FROM OR010300 WITH (NOLOCK) "
            'MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')"

            MySQLStr = "SELECT COUNT(OR01001) AS CC "
            MySQLStr = MySQLStr & "FROM (SELECT OR01001 "
            MySQLStr = MySQLStr & "FROM OR010300 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "') "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "Select OR20001 "
            MySQLStr = MySQLStr & "FROM OR200300 "
            MySQLStr = MySQLStr & "WHERE (OR20001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')) AS View_5 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            MyTID = Declarations.MyRec.Fields("CC").Value
        Loop While MyTID <> 0
        GetNewPRDID = MyID
    End Function

    Public Function GetNextPRDID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение следующего ID заказа на продажу 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyID As Double

        MySQLStr = "Select SY68002 "
        MySQLStr = MySQLStr & "FROM SY6803XX "
        MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyID = 1
            trycloseMyRec()
        Else
            MyID = Declarations.MyRec.Fields("SY68002").Value
            trycloseMyRec()
            MySQLStr = "UPDATE SY6803XX "
            MySQLStr = MySQLStr & "SET SY68002 = " & CStr(MyID + 1) & " "
            MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
            Declarations.MyConn.Execute(MySQLStr)
        End If
        GetNextPRDID = MyID
    End Function

    Public Function GetExchangeRate(ByVal MyCurr As Integer, ByVal Mydate As DateTime) As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение курса обмена валюты
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT SYCH006 "
        MySQLStr = MySQLStr & "FROM SYCH0100 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = " & CStr(MyCurr) & ") AND "
        MySQLStr = MySQLStr & "(SYCH004 <= CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(MM, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(yyyy, CONVERT(datetime,'" & Mydate & "',103))), 103)) AND "
        MySQLStr = MySQLStr & "(SYCH005 > CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(MM, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(yyyy, CONVERT(datetime,'" & Mydate & "',103))), 103)) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            GetExchangeRate = 1
        Else
            GetExchangeRate = Declarations.MyRec.Fields("SYCH006").Value
        End If
        trycloseMyRec()

    End Function

    Public Function ImportDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// импорт строк заказа из Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySQLStr As String
        Dim MyExcelCounter As Double                  'счетчик строк Excel
        Dim MyOrderCounter As Double                  'счетчик строк в заказе
        Dim MyItemCounter As Double                   'счетчик кол - ва Скальских кодов товара соотв. коду товара поставщика

        Dim MySuppItemCode As String                  'код товара поставщика
        Dim MyItemName As String                      'название товара
        Dim MyUOM As Integer                          'код единицы измерения
        Dim MyQTY As Double                           'кол - во
        Dim MyPrice As Double                         'цена за 1
        Dim MyPriCost As Double                       'расчетная себестоимость
        Dim c As Object
        Dim MyWeekQTY As Double                       'срок поставки
        Dim MyRez As Object

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            '---удаление старых значений из таблицы (для данного заказа)
            MyRez = MsgBox("Удалить старые данные из заказа?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                MySQLStr = "DELETE FROM  tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "')  "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                MyOrderCounter = 1
            Else
                MySQLStr = "SELECT MAX(OR03002) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MyOrderCounter = 1
                Else
                    If IsDBNull(Declarations.MyRec.Fields("CC").Value) = True Then
                        MyOrderCounter = 1
                    Else
                        MyOrderCounter = CInt(Declarations.MyRec.Fields("CC").Value) / 10 + 1
                    End If
                End If
            End If

            MyExcelCounter = 11

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing
                '------код запаса
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    Declarations.MyItemCode = ""
                Else
                    Declarations.MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                End If
                '------код запаса поставщика
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                End If
                If Len(MySuppItemCode) > 32 Then
                    MyRez = MsgBox("Ячейка C" & MyExcelCounter & " код товара поставщика не должен превышать 32 знакa. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" Then
                        MsgBox("Строка " & MyExcelCounter & " Обязательно должен быть занесен или код товара Scala, или код товара поставщика!", MsgBoxStyle.Critical, "Внимание!")
                    Else
                        If Trim(Declarations.MyItemCode) = "" Then
                            MySQLStr = "SELECT COUNT(*) AS CC "
                            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                trycloseMyRec()
                                MsgBox("Невозможно получить информацию о запасах из Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                                Exit While
                            Else
                                Declarations.MyRec.MoveFirst()
                                MyItemCounter = Declarations.MyRec.Fields("CC").Value
                                trycloseMyRec()
                            End If
                            If MyItemCounter = 0 Then
                                '---Запаса с таким кодом товара поставщика в Scala нет 
                                Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                            ElseIf MyItemCounter = 1 Then
                                '---Запас с таким кодом товара поставщика в Scala только один
                                MySQLStr = "SELECT SC01001 AS CC "
                                MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                                MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                    trycloseMyRec()
                                    Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                                Else
                                    Declarations.MyRec.MoveFirst()
                                    Declarations.MyItemCode = Declarations.MyRec.Fields("CC").Value
                                    trycloseMyRec()
                                End If
                            Else
                                '---В Scala несколько запасов с таким кодом товара поставщика 
                                MySelectItemBySuppCode = New SelectItemBySuppCode
                                MySelectItemBySuppCode.MyItemSuppCode = Trim(MySuppItemCode)
                                MySelectItemBySuppCode.MyWindowFrom = "Import"
                                MySelectItemBySuppCode.ShowDialog()
                            End If
                        End If
                    End If



                    If (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) Is Double) Then
                        MyRez = MsgBox("Ячейка D" & MyExcelCounter & " наименование товара должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    Else
                        If Len(appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) > 50 Then
                            MyRez = MsgBox("Ячейка D" & MyExcelCounter & " наименование товара не должно превышать 50 знаков. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                            If MyRez = vbYes Then
                                Exit While
                            End If
                        Else
                            '------название запаса
                            MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value
                            If (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value) Is Double) Then
                                MyRez = MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                                If c.Text = "" Then
                                    MyRez = MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' не выбрано из списка формы. Необходимо выбрать значение из выпадающего списка. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                    If MyRez = vbYes Then
                                        Exit While
                                    End If
                                Else
                                    '------код единицы измерения
                                    MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                                    If (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                        MyRez = MsgBox("Ячейка F" & MyExcelCounter & " значение 'Количество' должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                        If MyRez = vbYes Then
                                            Exit While
                                        End If
                                    Else
                                        If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                            MyRez = MsgBox("Ячейка F" & MyExcelCounter & " значение 'Количество' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                            If MyRez = vbYes Then
                                                Exit While
                                            End If
                                        Else
                                            '----------кол - во в заказе
                                            MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                                            If (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                                MyRez = MsgBox("Ячейка G" & MyExcelCounter & " значение 'Цена без НДС' должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                                If MyRez = vbYes Then
                                                    Exit While
                                                End If
                                            Else
                                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                                    MyRez = MsgBox("Ячейка G" & MyExcelCounter & " значение 'Цена без НДС' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                                    If MyRez = vbYes Then
                                                        Exit While
                                                    End If
                                                Else
                                                    '----------Цена
                                                    MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                                                    '----------Срок поставки
                                                    Try
                                                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                                                        If MyWeekQTY >= 0 Then
                                                            '----------Расчетная себестоимость
                                                            Try
                                                                MyPriCost = Math.Round(appXLSRC.Worksheets(1).Range("J" & MyExcelCounter).Value, 2)
                                                            Catch ex As Exception
                                                                MyPriCost = 0
                                                            End Try
                                                            AddRow(Declarations.MyOrderNum, Right("000000" & CStr(MyOrderCounter * 10), 6), Declarations.MyItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY, Trim(MySuppItemCode), MyPriCost)

                                                            MyOrderCounter = MyOrderCounter + 1.0
                                                        Else
                                                            MyRez = MsgBox("Ячейка I" & MyExcelCounter & " Срок поставки внесен некорректно. Должно быть число большее или равное нулю. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                                            If MyRez = vbYes Then
                                                                Exit While
                                                            End If
                                                        End If
                                                    Catch ex As Exception
                                                        MyRez = MsgBox("Ячейка I" & MyExcelCounter & " Срок поставки внесен некорректно. Должно быть число большее или равное нулю. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                                        If MyRez = vbYes Then
                                                            Exit While
                                                        End If
                                                    End Try
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------удаление объектов
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("Процедура импорта строк заказа на продажу завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Function ImportDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// импорт строк заказа из Libre Office
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyExcelCounter As Double                  'счетчик строк Excel
        Dim MyOrderCounter As Double                  'счетчик строк в заказе
        Dim MyItemCounter As Double                   'счетчик кол - ва Скальских кодов товара соотв. коду товара поставщика
        Dim MySuppItemCode As String                  'код товара поставщика
        Dim MyItemName As String                      'название товара
        Dim MyUOMStr As String                        'единица измерения
        Dim MyUOM As Integer                          'код единицы измерения
        Dim MyQTY As Double                           'кол - во
        Dim MyPrice As Double                         'цена за 1
        Dim MyPriCost As Double                       'расчетная себестоимость
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       'Версия документа
        Dim MySQLStr As String                        'SQL запрос
        Dim MyRez As Object
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

        '---Проверяем версию листа Excel
        MyVersion = oSheet.getCellRangeByName("A1").String
        If MyVersion = "" Then
            MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
            oWorkBook.Close(True)
            Exit Function
        Else
            MySQLStr = "SELECT Version "
            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
            MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                trycloseMyRec()
                oWorkBook.Close(True)
                Exit Function
            Else
                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                    trycloseMyRec()
                Else
                    MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "Внимание!")
                    trycloseMyRec()
                    oWorkBook.Close(True)
                    Exit Function
                End If
            End If
        End If

        '---удаление старых значений из таблицы (для данного заказа)
        MyRez = MsgBox("Удалить старые данные из заказа?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = vbYes Then
            MySQLStr = "DELETE FROM  tbl_OR030300 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "')  "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MyOrderCounter = 1
        Else
            MySQLStr = "SELECT MAX(OR03002) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MyOrderCounter = 1
            Else
                If IsDBNull(Declarations.MyRec.Fields("CC").Value) = True Then
                    MyOrderCounter = 1
                Else
                    MyOrderCounter = CInt(Declarations.MyRec.Fields("CC").Value) / 10 + 1
                End If
            End If
        End If

        MyExcelCounter = 11

        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------код запаса
            Declarations.MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            '------код запаса поставщика
            MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
            If Len(MySuppItemCode) > 32 Then
                MyRez = MsgBox("Ячейка C" & MyExcelCounter & " код товара поставщика не должен превышать 32 знакa. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = vbYes Then
                    Exit While
                End If
            Else
                If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" Then
                    MsgBox("Строка " & MyExcelCounter & " Обязательно должен быть занесен или код товара Scala, или код товара поставщика!", MsgBoxStyle.Critical, "Внимание!")
                Else
                    If Trim(Declarations.MyItemCode) = "" Then
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("Невозможно получить информацию о запасах из Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyItemCounter = Declarations.MyRec.Fields("CC").Value
                            trycloseMyRec()
                        End If
                        If MyItemCounter = 0 Then
                            '---Запаса с таким кодом товара поставщика в Scala нет 
                            Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                        ElseIf MyItemCounter = 1 Then
                            '---Запас с таким кодом товара поставщика в Scala только один
                            MySQLStr = "SELECT SC01001 AS CC "
                            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                trycloseMyRec()
                                Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                            Else
                                Declarations.MyRec.MoveFirst()
                                Declarations.MyItemCode = Declarations.MyRec.Fields("CC").Value
                                trycloseMyRec()
                            End If
                        Else
                            '---В Scala несколько запасов с таким кодом товара поставщика 
                            MySelectItemBySuppCode = New SelectItemBySuppCode
                            MySelectItemBySuppCode.MyItemSuppCode = Trim(MySuppItemCode)
                            MySelectItemBySuppCode.MyWindowFrom = "Import"
                            MySelectItemBySuppCode.ShowDialog()
                        End If
                    End If
                End If
            End If

            If oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") Then
                MyRez = MsgBox("Ячейка D" & MyExcelCounter & " наименование товара должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = vbYes Then
                    Exit While
                End If
            Else
                If Len(Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)) > 50 Then
                    MyRez = MsgBox("Ячейка D" & MyExcelCounter & " наименование товара не должно превышать 50 знаков. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    '------название запаса
                    MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
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
                    '----------Расчетная себестоимость
                    Try
                        MyPriCost = oSheet.getCellRangeByName("J" & MyExcelCounter).Value
                    Catch ex As Exception
                        MyPriCost = 0
                    End Try

                    AddRow(Declarations.MyOrderNum, Right("000000" & CStr(MyOrderCounter * 10), 6), Declarations.MyItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY, Trim(MySuppItemCode), MyPriCost)
                    MyOrderCounter = MyOrderCounter + 1.0
                End If
            End If
            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("Процедура импорта строк заказа на продажу завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Function ImportRequestDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// импорт строк запроса на поиск из Excel (спецификации)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySQLStr As String
        Dim MyExcelCounter As Double                  'счетчик строк Excel
        Dim MyItemCounter As Double                   'счетчик кол - ва Скальских кодов товара соотв. коду товара поставщика

        Dim MySuppItemCode As String                  'код товара поставщика
        Dim MyScalaSuppItemCode As String             'код товара поставщика в Scala
        Dim MyItemName As String                      'название товара
        Dim MyScalaItemName As String                 'название товара в Scala
        Dim MyUOM As Integer                          'код единицы измерения
        Dim MyScalaUOM As Integer                     'код единицы измерения в Scala
        Dim MyQTY As Double                           'кол - во
        Dim MyPrice As Double                         'цена за 1
        Dim c As Object
        Dim MyWeekQTY As Double                       'срок поставки
        Dim MyRez As Object

        MyScalaSuppItemCode = ""
        MyScalaItemName = ""
        MyScalaUOM = -1

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            '---удаление старых значений из таблицы (для данного заказа)
            MyRez = MsgBox("Удалить старые данные из запроса?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                MySQLStr = "DELETE FROM  tbl_SupplSearchItems "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = N'" & Declarations.MyRequestNum & "')  "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            MyExcelCounter = 11

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Or _
                Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Or _
                Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing

                '------код запаса Scala
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    Declarations.MyItemCode = ""
                Else
                    Declarations.MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                    '---проверка что такой код есть в Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        trycloseMyRec()
                        MyItemCounter = 0
                        'MsgBox("Ячейка B" & MyExcelCounter & ". Невозможно получить информацию о запасах из Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        'Exit While
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyItemCounter = Declarations.MyRec.Fields("CC").Value
                        trycloseMyRec()
                    End If
                    If MyItemCounter = 0 Then
                        'MsgBox("Ячейка B" & MyExcelCounter & ". такого кода товара в Scala нет.", MsgBoxStyle.Critical, "Внимание!")
                        'Exit While
                        Declarations.MyItemCode = ""
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
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    If Trim(MyItemCode) = "" Then
                        MySuppItemCode = ""
                    Else
                        MySuppItemCode = MyScalaSuppItemCode
                    End If
                Else
                    If Trim(MyItemCode) = "" Then
                        MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                    Else
                        MySuppItemCode = MyScalaSuppItemCode
                    End If
                    ''------Проверка уникальности кода товара поставщика в запросе
                    'MySQLStr = "SELECT COUNT(ID) AS CC "
                    'MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                    'MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    'MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
                    'InitMyConn(False)
                    'InitMyRec(False, MySQLStr)
                    'If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    '    trycloseMyRec()
                    '    MsgBox("Ячейка C" & MyExcelCounter & " Ошибка проверки уникальности кода товара производителя в запросе. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    '    Exit While
                    'Else
                    '    If Declarations.MyRec.Fields("CC").Value = 0 Then
                    '        trycloseMyRec()
                    '    Else
                    '        trycloseMyRec()
                    '        MsgBox("Ячейка C" & MyExcelCounter & " Ошибка - товар с кодом товара производителя " & Trim(MySuppItemCode) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
                    '        Exit While
                    '    End If
                    'End If
                End If


                '------Название запаса
                If (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) Is Double) Then
                    If Trim(MyItemCode) = "" Then
                        MyItemName = ""
                    Else
                        MyItemName = MyScalaItemName
                    End If
                Else
                    If Len(appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) > 254 Then
                        MsgBox("Ячейка D" & MyExcelCounter & " наименование товара не должно превышать 254 знака. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                        Exit While
                    Else
                        If Trim(MyItemCode) = "" Then
                            MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value
                        Else
                            MyItemName = MyScalaItemName
                        End If
                        ''------Проверка уникальности названия товара в запросе
                        'MySQLStr = "SELECT COUNT(ID) AS CC "
                        'MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                        'MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                        'MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
                        'InitMyConn(False)
                        'InitMyRec(False, MySQLStr)
                        'If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        '    trycloseMyRec()
                        '    MsgBox("Ячейка D" & MyExcelCounter & " Ошибка проверки уникальности названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                        '    Exit While
                        'Else
                        '    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        '        trycloseMyRec()
                        '    Else
                        '        trycloseMyRec()
                        '        MsgBox("Ячейка C" & MyExcelCounter & " Ошибка - товар с названием " & Trim(MyItemName) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
                        '        Exit While
                        '    End If
                        'End If
                    End If
                End If

                '------Проверка уникальности кода товара поставщика + названия товара в запросе
                MySQLStr = "SELECT COUNT(ID) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
                MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("Ячейка D" & MyExcelCounter & " Ошибка проверки уникальности кода товара производителя + названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                    Exit While
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                    Else
                        trycloseMyRec()
                        MsgBox("Ячейка C" & MyExcelCounter & " Ошибка - товар с кодом производителя " & Trim(MySuppItemCode) & " и названием " & Trim(MyItemName) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
                        Exit While
                    End If
                End If


                '-------Проверка заполнения первых 3 - х полей
                'If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                '-----код Scala не заносим
                If Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                    MsgBox("Строка " & MyExcelCounter & " Обязательно должен быть занесен код товара поставщика или название товара!", MsgBoxStyle.Critical, "Внимание!")
                    Exit While
                Else
                    '------код единицы измерения
                    If (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value) Is Double) Then
                        If Trim(MyItemCode) = "" Then
                            MyRez = MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                            If MyRez = vbYes Then
                                Exit While
                            End If
                        Else
                            MyUOM = MyScalaUOM
                        End If
                    Else
                        c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                        If c.Text = "" Then
                            If Trim(MyItemCode) = "" Then
                                MyRez = MsgBox("Ячейка E" & MyExcelCounter & " значение 'Единица измерения' не выбрано из списка формы. Необходимо выбрать значение из выпадающего списка. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                MyUOM = MyScalaUOM
                            End If
                        Else
                            If Trim(MyItemCode) = "" Then
                                MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                            Else
                                MyUOM = MyScalaUOM
                            End If
                            '----------кол - во в запросе
                            If (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                MyRez = MsgBox("Ячейка F" & MyExcelCounter & " значение 'Количество' должно быть заполнено. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                    MyRez = MsgBox("Ячейка F" & MyExcelCounter & " значение 'Количество' должно быть заполнено числовым значением. Прервать выполнение импорта?", MsgBoxStyle.YesNo, "Внимание!")
                                    If MyRez = vbYes Then
                                        Exit While
                                    End If
                                Else
                                    MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                                    '----------Цена
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                        MyPrice = 0
                                    Else
                                        MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                                    End If
                                    '----------Срок поставки
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value) Is Double) Then
                                        MyWeekQTY = 0
                                    Else
                                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                                    End If
                                    '----------Занесение строки
                                    'AddRequestRow(Declarations.MyRequestNum, Declarations.MyItemCode, MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
                                    '---код Scala не заносим
                                    AddRequestRow(Declarations.MyRequestNum, "", MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
                                End If
                            End If
                        End If
                    End If
                End If
                MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------удаление объектов
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("Процедура импорта строк запроса на поиск поставщика завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Function ImportRequestDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// импорт строк запроса на поиск из LibreOffice (спецификации)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
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

        '---Проверяем версию листа Excel
        MyVersion = oSheet.getCellRangeByName("A1").String
        If MyVersion = "" Then
            MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
            oWorkBook.Close(True)
            Exit Function
        Else
            MySQLStr = "SELECT Version "
            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
            MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                trycloseMyRec()
                oWorkBook.Close(True)
                Exit Function
            Else
                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                    trycloseMyRec()
                Else
                    MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "Внимание!")
                    trycloseMyRec()
                    oWorkBook.Close(True)
                    Exit Function
                End If
            End If
        End If

        '---удаление старых значений из таблицы (для данного заказа)
        MyRez = MsgBox("Удалить старые данные из запроса?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = vbYes Then
            MySQLStr = "DELETE FROM  tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (SupplSearchID = N'" & Declarations.MyRequestNum & "')  "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        MyExcelCounter = 11
        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------код запаса Scala
            Declarations.MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            If Declarations.MyItemCode.Equals("") = False Then
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
                    Declarations.MyItemCode = ""
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
            If Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String).Equals("") Then
                If Trim(MyItemCode) = "" Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = MyScalaSuppItemCode
                End If
            Else
                If Trim(MyItemCode) = "" Then
                    MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
                Else
                    MySuppItemCode = MyScalaSuppItemCode
                End If
            End If

            '------Название запаса
            If Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String).Equals("") Then
                If Trim(MyItemCode) = "" Then
                    MyItemName = ""
                Else
                    MyItemName = MyScalaItemName
                End If
            Else
                If Len(Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)) > 254 Then
                    MsgBox("Ячейка D" & MyExcelCounter & " наименование товара не должно превышать 254 знака.", MsgBoxStyle.YesNo, "Внимание!")
                    Exit While
                Else
                    If Trim(MyItemCode) = "" Then
                        MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
                    Else
                        MyItemName = MyScalaItemName
                    End If
                End If
            End If

            '------Проверка уникальности кода товара поставщика + названия товара в запросе
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
            MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
            MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("Ячейка D" & MyExcelCounter & " Ошибка проверки уникальности кода товара производителя + названия товара. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                Exit While
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                Else
                    trycloseMyRec()
                    MsgBox("Ячейка C" & MyExcelCounter & " Ошибка - товар с кодом производителя " & Trim(MySuppItemCode) & " и названием " & Trim(MyItemName) & " уже присутствует в запросе на поиск поставщика.", MsgBoxStyle.Critical, "Внимание!")
                    Exit While
                End If
            End If

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
                '----------Занесение строки
                '---код Scala не заносим
                AddRequestRow(Declarations.MyRequestNum, "", MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
            End If
            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("Процедура импорта строк запроса на поиск поставщика завершена.", vbOKOnly, "Внимание!")
    End Function

    Public Sub AddRequestRow(ByVal MyRequest As Integer, ByVal MyItemCode As String, ByVal MySuppItemCode As String, ByVal MyItemName As String, _
        ByVal MyUOM As Integer, ByVal MyQTY As Double, ByVal MyPrice As Double, ByVal MyWeekQTY As Double)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание строки в запросе на поиск поставщика
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_SupplSearchItems"
        MySQLStr = MySQLStr & "(SupplSearchID, ItemID, ItemSuppID, ItemName, UOM, QTY, LeadTimeWeek, Comments) "
        MySQLStr = MySQLStr & "VALUES ("
        MySQLStr = MySQLStr & MyRequest.ToString & ", "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MyItemCode), "'", "''") & "', "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MySuppItemCode), "'", "''") & "', "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MyItemName), "'", "''") & "', "
        MySQLStr = MySQLStr & MyUOM.ToString & ", "
        MySQLStr = MySQLStr & Replace(MyQty.ToString, ",", ".") & ", "
        If MyWeekQTY = 0 Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & Replace(MyWeekQTY.ToString, ",", ".") & ", "
        End If
        MySQLStr = MySQLStr & "N'' "
        MySQLStr = MySQLStr & ") "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

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

    Public Function AddRow(ByVal MyOrder As String, ByVal MyStr As String, ByVal MyItemCode As String, ByVal MyItemName As String, _
        ByVal MyUOM As Integer, ByVal MyQTY As Double, ByVal MyPrice As Double, ByVal MyWeekQTY As Double, _
        ByVal MySuppItemCode As String, ByVal MyPriCost As Double)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание строки заказа на продажу
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        Dim MySQLStr As String

        cmd.ActiveConnection = Declarations.MyConn
        cmd.CommandText = "spp_SalesWorkplace4_ImportRow"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.CommandTimeout = 300

        MyParam = cmd.CreateParameter("@MyOrderNum", 129, ADODB.ParameterDirectionEnum.adParamInput, 10)
        cmd.Parameters.Append(MyParam)
        MyParam.Value = MyOrder

        MyParam1 = cmd.CreateParameter("@MyStrNum", 129, ADODB.ParameterDirectionEnum.adParamInput, 6)
        cmd.Parameters.Append(MyParam1)
        MyParam1.Value = MyStr

        MyParam2 = cmd.CreateParameter("@MyItemCode", 129, ADODB.ParameterDirectionEnum.adParamInput, 35)
        cmd.Parameters.Append(MyParam2)
        MyParam2.Value = MyItemCode

        MyParam3 = cmd.CreateParameter("@MyItemName", 129, ADODB.ParameterDirectionEnum.adParamInput, 51)
        cmd.Parameters.Append(MyParam3)
        MyParam3.Value = MyItemName

        MyParam4 = cmd.CreateParameter("@MyUOM", 3, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam4)
        MyParam4.Value = MyUOM

        MyParam5 = cmd.CreateParameter("@MyQTY", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam5)
        MyParam5.Value = MyQTY

        MyParam6 = cmd.CreateParameter("@MyPrice", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam6)
        MyParam6.Value = MyPrice

        MyParam7 = cmd.CreateParameter("@MyWeekQTY", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam7)
        MyParam7.Value = MyWeekQTY

        MyParam8 = cmd.CreateParameter("@MyPriCost", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam8)
        MyParam8.Value = MyPriCost

        cmd.Execute()

        '-----обработка сроков поставки
        'MySQLStr = "UPDATE tbl_OR030300 "
        'MySQLStr = MySQLStr & "SET WeekQTY = " & Replace(CStr(MyWeekQTY), ",", ".") & " "
        'MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
        'MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
        'InitMyConn(False)
        'Declarations.MyConn.Execute(MySQLStr)

        '-----Добавление кода товара поставщика
        MySQLStr = "UPDATE tbl_OR030300 "
        MySQLStr = MySQLStr & "SET SuppItemCode = N'" & MySuppItemCode & "' "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
        MySQLStr = MySQLStr & "AND (OR03002 = N'" & MyStr & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

    End Function

    Public Function CheckRights(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Проверка прав пользователя - является ли членом группы CRMManagers
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyCCPermission = False
            CheckRights = "Запрещено"
        Else
            Declarations.MyCCPermission = True
            CheckRights = "Разрешено"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 5")
        Declarations.MyCCPermission = False
        CheckRights = "Запрещено"
    End Function

    Public Function CheckRights1(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Проверка прав пользователя - является ли членом группы CRMDirector
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyPermission = False
            CheckRights1 = "Запрещено"
        Else
            Declarations.MyPermission = True
            CheckRights1 = "Разрешено"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 6")
        Declarations.MyPermission = False
        CheckRights1 = "Запрещено"
    End Function

    Public Function CheckRights2(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Проверка прав пользователя - является ли членом группы ProposalManager
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyCPPermission = False
            CheckRights2 = "Запрещено"
        Else
            Declarations.MyCPPermission = True
            CheckRights2 = "Разрешено"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 7")
        Declarations.MyCPPermission = False
        CheckRights2 = "Запрещено"
    End Function

    Public Sub SendAddInfoReminder(ByVal MyItemCode As String, ByVal MySalesman As String, ByVal MyType As Integer)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка по почте уведомления отделу закупок о неверных картинках, названиях и описаниях
        '// MyType 0 - картинка 1 - название 2 - описание
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_SalesWorkplace4_SendReminder N'" & Trim(MyItemCode) & "', N'" & MySalesman & "', " & CStr(MyType)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub SendInfoByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewRequestState As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка по почте уведомления поисковикам об изменениях в запросах
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
        MyMsgStr = MyMsgStr & "Новое состояние запроса: " & NewRequestState & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "С уважением," & Chr(13)
        MyMsgStr = MyMsgStr + "ООО ""Скандика"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. На письмо просьба не отвечать, это автоматическая рассылка. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

    Public Sub SendCommentByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewComment As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка по почте нового комментария поисковикам
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
        msg.Subject = "Уведомление о новом комментарии запроса на поиск"
        MyMsgStr = "Уважаемый коллега!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "Добавлен новый комментарий к следующему запросу на поиск:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " от " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "Для клиента: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "Продавец: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "Добавлен комментарий: " & NewComment & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "С уважением," & Chr(13)
        MyMsgStr = MyMsgStr + "ООО ""Скандика"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. На письмо просьба не отвечать, это автоматическая рассылка. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

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

    Public Function GetSrchManagerEmailFromDB() As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение Email менеджера поисковиков
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MySTR As String
        Dim MyFlag As Integer

        MySTR = ""
        MyFlag = 0
        MySQLStr = "SELECT RM.dbo.RM660100.RM66003 "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_Searchers INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON tbl_SupplSearch_Searchers.PurchID = RM.dbo.RM660100.RM66001 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_Searchers.IsLeader = 1) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            GetSrchManagerEmailFromDB = ""
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                If Trim(Declarations.MyRec.Fields("RM66003").Value.ToString) <> "" Then
                    If MyFlag <> 0 Then
                        MySTR = MySTR & ";"
                    End If
                    MySTR = MySTR & Trim(Declarations.MyRec.Fields("RM66003").Value.ToString)
                End If
                Declarations.MyRec.MoveNext()
                MyFlag = MyFlag + 1
            End While
            GetSrchManagerEmailFromDB = Trim(MySTR)
        End If
    End Function

    Public Function CheckSalesman(ByVal SalesmanCode As String, ByVal CustomerCode As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка что продавец должен быть из того же кост центра, что и продавец, за которым закреплен клиент
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustomerCC As String
        Dim SalesmanCC As String
        Dim ExclQTY As Integer

        '---проверка - не в исключениях ли продавец
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_Sales_Groups_CC_Exclude "
        MySQLStr = MySQLStr & "WHERE (Upper(UserName) = N'" & Declarations.UserCode.ToUpper() & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckSalesman = "Невозможно определить - добавлен продавец в исключения или нет "
            Exit Function
        Else
            ExclQTY = Declarations.MyRec.Fields("CC").Value
        End If

        If ExclQTY > 0 Then '---продавец в исключениях по проверке
            CheckSalesman = ""
            Exit Function
        End If

        If Trim(SalesmanCode) = "" Or Trim(CustomerCode) = "" Then
            CheckSalesman = ""
            Exit Function
        Else
            '--CC продавца
            MySQLStr = "SELECT SUBSTRING(ST01021, 7, 3) AS CC "
            MySQLStr = MySQLStr & "FROM ST010300 "
            MySQLStr = MySQLStr & "WHERE (ST01001 = N'" & Trim(SalesmanCode) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                SalesmanCC = "CC продавца не найден"
            Else
                SalesmanCC = Declarations.MyRec.Fields("CC").Value.ToString
                trycloseMyRec()
            End If

            '--CC клиента
            MySQLStr = "SELECT     SUBSTRING(ST010300.ST01021, 7, 3) AS CC "
            MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON SL010300.SL01035 = ST010300.ST01001 "
            MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(CustomerCode) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                CustomerCC = "CC клиента не найден"
            Else
                CustomerCC = Declarations.MyRec.Fields("CC").Value.ToString
                trycloseMyRec()
            End If

            '---Сверка
            If SalesmanCC <> CustomerCC And CustomerCC <> "CC клиента не найден" Then
                If My.Settings.CheckCC.ToUpper() = "ДА" Then
                    CheckSalesman = "Кост центр продавца должен быть таким же, как кост центр продавца, за которым закреплен клиент. "
                    CheckSalesman = CheckSalesman & "кост центр продавца: " & SalesmanCC & " кост центр клиента: " & CustomerCC
                Else
                    CheckSalesman = ""
                End If
                Exit Function
            Else
                CheckSalesman = ""
                Exit Function
            End If
        End If
    End Function
End Module
