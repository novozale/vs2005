Module OBOBettermanInvoiceFunctions

    Public Sub OpenOBOBettermanInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие файла с инвойсом OBO Betterman
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog
        Dim MyPurchOrder As String
        Dim MyInvoiceNumber As String
        Dim MyInvoiceDate As DateTime
        Dim MySQLStr As String                        'рабочая строка
        Dim i As Integer                              'счетчик строк

        '---получение имени файла
        MyDlg = New OpenFileDialog
        MyDlg.Filter = "Файлы Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        If MyDlg.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        '---попытка открытия документа
        Try
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        Catch ex As Exception
        End Try

        appXLSRC = CreateObject("Excel.Application")
        Try
            appXLSRC.Workbooks.Open(MyDlg.FileName)
        Catch ex As Exception
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("Ошибка " + ex.Message)
            'Exit Sub
        End Try

        '---Определение кода и названия поставщика
        '---сначала проверяем - может быть это ОБО Беттерман
        '---Номер заказа на закупку
        MyPurchOrder = OBOBettermanGetPurchOrderNum(appXLSRC.Worksheets(1).Range("B18").Value.ToString)
        If MyPurchOrder = "" Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("Не определен обобщенный номер заказа на закупку Электроскандии", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        End If

        '---Код поставщика
        MySQLStr = "Select SupplierCode AS Code "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & MyPurchOrder & "') AND "
        MySQLStr = MySQLStr & "(SupplierCode = N'1029') " '---АББ
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("Поставщик данной СФ или заказы на закупку, соответствующие данной СФ не найдены в Scala", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        Else
            Main.TextBox1.Text = declarations.MyRec.Fields("Code").Value
            trycloseMyRec()
        End If

        '---N СФ поставщика
        MyInvoiceNumber = OBOBettermanGetInvoiceNum(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        If MyInvoiceNumber = "" Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            MsgBox("Не определен номер счета фактуры поставщика ОБО Беттерман", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        Else
            Main.textBox3.Text = MyInvoiceNumber
        End If

        '---дата СФ поставщика
        MyInvoiceDate = OBOBettermanGetInvoiceDate(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        If MyInvoiceDate = CDate("31/12/9999") Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            Main.textBox3.Text = ""
            MsgBox("Не определена дата счета фактуры поставщика ОБО Беттерман", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        Else
            Main.textBox4.Text = Format(MyInvoiceDate, "dd/MM/yyyy")
        End If

        '---Валюта СФ поставщика
        If InStr(UCase(appXLSRC.Worksheets(1).Range("B13").Value.ToString), "РУБЛЬ") > 0 Then
            Main.textBox5.Text = 0
        Else
            Main.textBox5.Text = 12
        End If

        '---Проверяем - может быть эта СФ уже прогружена / введена
        MySQLStr = "SELECT COUNT(PC190300.PC19001) AS CC "
        MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC190300.PC19001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC190300.PC19012 = N'" & Main.textBox3.Text & "') AND "
        MySQLStr = MySQLStr & "(PC010300.PC01052 = N'" & MyPurchOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.Fields("CC").Value > 0 Then
            trycloseMyRec()
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            Main.textBox3.Text = ""
            Main.textBox4.Text = ""
            Main.textBox5.Text = ""
            Main.label6.Text = "данная СФ уже загружена в Scala"
            Exit Sub
        Else
            trycloseMyRec()
            Main.button3.Enabled = True
            Main.label6.Text = ""
        End If

        '---Выставляем значениия прогресс бара
        i = 22
        While Not Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString) = "Всего к оплате"
            i = i + 1
        End While
        Main.progressBar1.Minimum = 0
        Main.progressBar1.Maximum = i - 23
    End Sub

    Public Sub UploadOBOBettermanInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка файла с инвойсом OBO Betterman
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String

        MyRezStr = ""

        LoadOBOBettermanInvoiceToTMPTable()
        MyRezStr = CheckUOMInOrders()
        If MyRezStr = "" Then
            If CheckEmptyInOrders() = True Then
                MyRezStr = LoadInvoiceFromTMPTable()
            Else
                MsgBox("Проверьте присланый файл с СФ. Не указан либо код товара поставщика, либо страна, либо количество равно нулю (незаполнено), либо сумма без НДС за строку равна нулю (незаполнена).", MsgBoxStyle.Critical, "Внимание!")
            End If
        End If
            UploadingRezult(MyRezStr)
    End Sub

    Private Sub LoadOBOBettermanInvoiceToTMPTable()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка файла с инвойсом OBO Betterman во временную таблицу
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim i As Integer                                'счетчик
        Dim MyInvoice As String                         '--номер СФ
        Dim MyInvoiceDate As String                     '--дата СФ
        Dim MyInvoiceCurrCode As Integer                '--Код валюты СФ
        Dim MySalesmanCode As String                    '--код продавца
        Dim MySalesmanName As String                    '--имя продавца
        Dim MyInvoiceCurrExchRateStr As String          '--Курс валюты в инвойсе (строка)
        Dim MyInvoiceCurrExchRate As Double             '--Курс валюты в инвойсе
        Dim MyConsPurchaseOrderNum As String            '--Номер консолидированного заказа на закупку
        Dim MySupplierItemCode As String                '--код товара поставщика
        Dim MyQTY As Double                             '--количество
        Dim MySummWithoutVAT As Double                  '--Сумма без НДС за строку
        Dim MyCountryCode As String                     '-- Код страны производителя
        Dim MyCountry As String                         '-- страна производителя
        Dim MyGTD As String                             '-- ГТД

        '---Удаление старой временной таблицы
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyInvoice') "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyInvoice "
        InitMyConn(False)
        declarations.MyConn.Execute(MySQLStr)

        '---Создание новой временной таблицы
        MySQLStr = "CREATE TABLE #_MyInvoice( "
        MySQLStr = MySQLStr & "[ID] int, "                                 '--ID строки
        MySQLStr = MySQLStr & "[Invoice] [nvarchar](35), "                 '--номер СФ
        MySQLStr = MySQLStr & "[InvoiceDate] [datetime], "                 '--дата СФ
        MySQLStr = MySQLStr & "[InvoiceCurrCode] int, "                    '--Код валюты СФ
        MySQLStr = MySQLStr & "[SalesmanCode] [nvarchar](3), "             '--код продавца
        MySQLStr = MySQLStr & "[SalesmanName] [nvarchar](25), "            '--имя продавца
        MySQLStr = MySQLStr & "[InvoiceCurrExchRate] float, "              '--Курс валюты в инвойсе
        MySQLStr = MySQLStr & "[ConsPurchaseOrderNum] [nvarchar](10), "    '--Номер консолидированного заказа на закупку
        MySQLStr = MySQLStr & "[SupplierItemCode] [nvarchar](35), "        '--код товара поставщика
        MySQLStr = MySQLStr & "[QTY] float, "                              '--количество
        MySQLStr = MySQLStr & "[SummWithoutVAT] float, "                   '--Сумма без НДС за строку
        MySQLStr = MySQLStr & "[Country] nvarchar(50), "                   '-- страна производителя
        MySQLStr = MySQLStr & "[GTD] nvarchar (255), "                     '-- ГТД
        MySQLStr = MySQLStr & "[RestQTY] float  "                          '--Остаток - непринятое количество
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        declarations.MyConn.Execute(MySQLStr)

        MyInvoice = OBOBettermanGetInvoiceNum(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        MyInvoiceDate = CStr(Format(OBOBettermanGetInvoiceDate(appXLSRC.Worksheets(1).Range("B2").Value), "dd/MM/yyyy"))
        MyInvoiceCurrCode = CInt(Main.textBox5.Text)
        MySalesmanCode = declarations.SalesmanCode
        MySalesmanName = declarations.SalesmanName
        If MyInvoiceCurrCode = 0 Then
            MyInvoiceCurrExchRate = 1
        Else
            MyInvoiceCurrExchRateStr = OBOBettermanGetInvoiceCurrExchRate(appXLSRC.Worksheets(1).Range("B14").Value.ToString)
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MyInvoiceCurrExchRate = CDbl(MyInvoiceCurrExchRateStr)
            Else
                MyInvoiceCurrExchRate = CDbl(Replace(MyInvoiceCurrExchRateStr, ",", "."))
            End If
        End If
        
        MyConsPurchaseOrderNum = OBOBettermanGetPurchOrderNum(appXLSRC.Worksheets(1).Range("B18").Value.ToString)

        i = 23
        While Not Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value) = "Всего к оплате"
            MySupplierItemCode = OBOBettermanGetSupplierItemCode(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value)
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MyQTY = CDbl(Replace(appXLSRC.Worksheets(1).Range("K" & CStr(i)).Value.ToString, ".", ","))
            Else
                MyQTY = CDbl(appXLSRC.Worksheets(1).Range("K" & CStr(i)).Value.ToString)
            End If
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MySummWithoutVAT = CDbl(Replace(appXLSRC.Worksheets(1).Range("O" & CStr(i)).Value.ToString, ".", ","))
            Else
                MySummWithoutVAT = CDbl(appXLSRC.Worksheets(1).Range("O" & CStr(i)).Value.ToString)
            End If
            If appXLSRC.Worksheets(1).Range("Y" & CStr(i)).Value = Nothing Then
                MyCountryCode = "643"
            Else
                MyCountryCode = appXLSRC.Worksheets(1).Range("Y" & CStr(i)).Value.ToString
            End If

            '---Находим страну по коду страны
            MySQLStr = "SELECT SY24003 "
            MySQLStr = MySQLStr & "FROM SY240300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SY24001 = N'BM') AND (SY24002 = N'" & Right("000" & MyCountryCode, 3) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                MsgBox("Ячейка Y" & CStr(i) & " Такой код страны в Scala не обнаружен.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Exit Sub
            Else
                MyCountry = declarations.MyRec.Fields("SY24003").Value
                trycloseMyRec()
            End If
            'MyCountry = appXLSRC.Worksheets(1).Range("X" & CStr(i)).Value.ToString

            If appXLSRC.Worksheets(1).Range("AA" & CStr(i)).Value = Nothing Then
                MyGTD = ""
            Else
                MyGTD = Trim(appXLSRC.Worksheets(1).Range("AA" & CStr(i)).Value.ToString)
            End If

            MySQLStr = "INSERT INTO #_MyInvoice "
            MySQLStr = MySQLStr & "(ID, Invoice, InvoiceDate, InvoiceCurrCode, SalesmanCode, SalesmanName, InvoiceCurrExchRate, "
            MySQLStr = MySQLStr & "ConsPurchaseOrderNum, SupplierItemCode, QTY, SummWithoutVAT, Country, GTD, RestQTY) "
            MySQLStr = MySQLStr & "VALUES (" & CStr(i - 21) & ", "
            MySQLStr = MySQLStr & "N'" & MyInvoice & "', "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & MyInvoiceDate & "', 103), "
            MySQLStr = MySQLStr & CStr(MyInvoiceCurrCode) & ", "
            MySQLStr = MySQLStr & "N'" & MySalesmanCode & "', "
            MySQLStr = MySQLStr & "N'" & MySalesmanName & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyInvoiceCurrExchRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "N'" & MyConsPurchaseOrderNum & "', "
            MySQLStr = MySQLStr & "N'" & MySupplierItemCode & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(MySummWithoutVAT), ",", ".") & ", "
            MySQLStr = MySQLStr & "N'" & MyCountry & "', "
            MySQLStr = MySQLStr & "N'" & MyGTD & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
            InitMyConn(False)
            declarations.MyConn.Execute(MySQLStr)

            i = i + 1
        End While

        'MySQLStr = "SELECT * FROM #_MyInvoice "
        'InitMyConn(False)
        'InitMyRec(False, MySQLStr)
        'While declarations.MyRec.EOF <> True
        ' MsgBox("ID:" & declarations.MyRec.Fields("ID").Value.ToString & Chr(13) & " Invoice:" & declarations.MyRec.Fields("Invoice").Value.ToString & Chr(13) & " InvoiceDate:" & declarations.MyRec.Fields("InvoiceDate").Value.ToString & Chr(13) & " InvoiceCurrCode:" & declarations.MyRec.Fields("InvoiceCurrCode").Value.ToString & Chr(13) & " SalesmanCode:" & declarations.MyRec.Fields("SalesmanCode").Value.ToString & Chr(13) & " SalesmanName:" & declarations.MyRec.Fields("SalesmanName").Value.ToString & Chr(13) & " InvoiceCurrExchRate:" & declarations.MyRec.Fields("InvoiceCurrExchRate").Value.ToString & Chr(13) & " ConsPurchaseOrderNum:" & declarations.MyRec.Fields("ConsPurchaseOrderNum").Value.ToString & Chr(13) & " SupplierItemCode:" & declarations.MyRec.Fields("SupplierItemCode").Value.ToString & Chr(13) & " QTY:" & declarations.MyRec.Fields("QTY").Value.ToString & Chr(13) & " SummWithoutVAT:" & declarations.MyRec.Fields("SummWithoutVAT").Value.ToString & Chr(13) & " Country:" & declarations.MyRec.Fields("Country").Value.ToString & Chr(13) & " GTD:" & declarations.MyRec.Fields("GTD").Value.ToString & Chr(13) & " RestQTY:" & declarations.MyRec.Fields("RestQTY").Value.ToString, MsgBoxStyle.Information, "Внимание!")
        'declarations.MyRec.MoveNext()
        'End While

    End Sub

    Private Function OBOBettermanGetPurchOrderNum(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение нашего номера заказа из строки СФ ОБО Беттерман
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        'Позиция подстроки

        'MyStr = Mid(MyStr, 9)
        'If Len(MyStr) < 9 Then
        '    OBOBettermanGetPurchOrderNum = ""
        'Else
        '    MyPos = InStr(MyStr, " ")
        '    If MyPos = 0 Then
        '        OBOBettermanGetPurchOrderNum = ""
        '    Else
        '        MyStr = Mid(MyStr, 1, MyPos - 1)
        '        OBOBettermanGetPurchOrderNum = Right("0000000000" & MyStr, 10)
        '    End If
        'End If
        '---уроды постоянно пишут по разному - то с N, то заявка, то заказ - привязываемся к тому, что наш код начинается с 07
        MyPos = InStr(MyStr, "07")
        If MyPos = 0 Then
            OBOBettermanGetPurchOrderNum = ""
        Else
            MyStr = Mid(MyStr, MyPos, 10)
            OBOBettermanGetPurchOrderNum = Right("0000000000" & MyStr, 10)
        End If
    End Function

    Private Function OBOBettermanGetInvoiceNum(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение номера СФ из строки СФ ОБО Беттерман
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        'Позиция подстроки

        MyStr = Mid(MyStr, 16)
        If Len(MyStr) < 1 Then
            OBOBettermanGetInvoiceNum = ""
        Else
            MyPos = InStr(MyStr, " ")
            If MyPos = 0 Then
                OBOBettermanGetInvoiceNum = ""
            Else
                MyStr = Mid(MyStr, 1, MyPos - 1)
                OBOBettermanGetInvoiceNum = MyStr
            End If
        End If
    End Function

    Private Function OBOBettermanGetInvoiceDate(ByVal MyStr As String) As DateTime
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение даты СФ из строки СФ ОБО Беттерман
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        'Позиция подстроки

        MyStr = Mid(MyStr, 16)
        If Len(MyStr) < 1 Then
            OBOBettermanGetInvoiceDate = CDate("31/12/9999")
        Else
            MyPos = InStr(MyStr, " ")
            If MyPos = 0 Then
                OBOBettermanGetInvoiceDate = CDate("31/12/9999")
            Else
                MyStr = Mid(MyStr, MyPos + 1)
                If Len(MyStr) < 1 Then
                    OBOBettermanGetInvoiceDate = CDate("31/12/9999")
                Else
                    MyPos = InStr(MyStr, " ")
                    If MyPos = 0 Then
                        OBOBettermanGetInvoiceDate = CDate("31/12/9999")
                    Else
                        MyStr = Mid(MyStr, MyPos + 1)
                        OBOBettermanGetInvoiceDate = CDate(MyStr)
                    End If
                End If
            End If
        End If
    End Function

    Private Function OBOBettermanGetInvoiceCurrExchRate(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение курса обмена валюты из строки СФ ОБО Беттерман
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        'Позиция подстроки

        MyPos = InStr(MyStr, "курс")
        If MyPos = 0 Then
            OBOBettermanGetInvoiceCurrExchRate = "0"
        Else
            MyStr = Mid(MyStr, MyPos + 5)
            OBOBettermanGetInvoiceCurrExchRate = MyStr
        End If
    End Function

    Private Function OBOBettermanGetSupplierItemCode(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение кода товара поставщика из строки СФ ОБО Беттерман
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        'Позиция подстроки
        Dim subStr As String

        MyPos = InStr(MyStr, "Код:")
        If MyPos = 0 Then
            OBOBettermanGetSupplierItemCode = Trim(MyStr)
        Else
            subStr = MyStr.Substring(MyPos + 4)
            OBOBettermanGetSupplierItemCode = Trim(subStr)
        End If
    End Function
End Module
