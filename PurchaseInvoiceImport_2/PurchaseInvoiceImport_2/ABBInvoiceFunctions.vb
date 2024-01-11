Imports System.IO

Module ABBInvoiceFunctions

    Public Sub OpenABBInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие файла с инвойсом ABB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog
        Dim MyPurchOrder As String
        Dim i As Integer                              'счетчик
        Dim MySupplierFound As Integer                'флаг - найден поставщик или нет
        Dim MyInvoiceUploaded As Integer              'флаг - уже прогружена СФ или нет
        Dim MySQLStr As String                        'рабочая строка

        '---получение имени файла
        MyDlg = New OpenFileDialog
        MyDlg.Filter = "Файлы XML (*.xml)|*.xml"
        If MyDlg.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        '---Исправление косяков АББ - удаление срмволов перевода строки-----------------------
        RemoveCarrigeReturnSym(MyDlg.FileName)

        '---попытка открытия документа
        MyDoc = New Xml.XmlDocument
        Try
            MyDoc.Load(MyDlg.FileName)
        Catch ex As Exception
            MsgBox("Ошибка " + ex.Message)
            Exit Sub
        End Try

        '---Определение кода и названия поставщика
        '---сначала проверяем - может быть это ABB
        Try
            MyHeaderNode = MyDoc.DocumentElement.ChildNodes(0)
            MyFirstItemNode = MyHeaderNode.ChildNodes(1)
            MyHeaderNode = MyHeaderNode.ChildNodes(0)
            MyItemNodeList = MyFirstItemNode.ChildNodes
            MySupplierFound = 0
            For i = 0 To MyItemNodeList.Count - 1
                'MyPurchOrder = Right("0000000000" + Right(MyItemNodeList(i).ChildNodes(13).InnerText, MyItemNodeList(i).ChildNodes(13).InnerText.Length - 3), 10)
                MyPurchOrder = Right("0000000000" + MyItemNodeList(i).ChildNodes(13).InnerText, 10)

                MySQLStr = "Select SupplierCode AS Code "
                MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (ID = N'" & MyPurchOrder & "') AND "
                MySQLStr = MySQLStr & "(SupplierCode = N'3046') " '---АББ
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                Else
                    Main.TextBox1.Text = declarations.MyRec.Fields("Code").Value
                    trycloseMyRec()
                    MySupplierFound = 1
                    Exit For
                End If
            Next i
            If MySupplierFound = 0 Then
                Main.label6.Text = ""
                Main.button3.Enabled = False
                Throw New Exception("Поставщик данной СФ или заказы на закупку, соответствующие данной СФ не найдены в Scala")
            Else
                '---заполняем оставшиеся поля формы
                Main.textBox3.Text = MyHeaderNode.ChildNodes(0).InnerText      'СФ поставщика
                Main.textBox4.Text = Replace(MyHeaderNode.ChildNodes(1).InnerText, ".", "/")
                If InStr(UCase(MyHeaderNode.ChildNodes(4).InnerText), "ЕВРО") > 0 Then
                    Main.textBox5.Text = 12
                Else
                    Main.textBox5.Text = 0
                End If
                '---Проверяем - может быть эта СФ уже прогружена / введена
                MyInvoiceUploaded = 0
                For i = 0 To MyItemNodeList.Count - 1
                    'MyPurchOrder = Right("0000000000" + Right(MyItemNodeList(i).ChildNodes(13).InnerText, MyItemNodeList(i).ChildNodes(13).InnerText.Length - 3), 10)
                    MyPurchOrder = Right("0000000000" + MyItemNodeList(i).ChildNodes(13).InnerText, 10)
                    MySQLStr = "SELECT COUNT(PC190300.PC19001) AS CC "
                    MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) INNER JOIN "
                    MySQLStr = MySQLStr & "PC010300 ON PC190300.PC19001 = PC010300.PC01001 "
                    MySQLStr = MySQLStr & "WHERE (PC190300.PC19012 = N'" & Main.textBox3.Text & "') AND "
                    MySQLStr = MySQLStr & "(PC010300.PC01052 = N'" & MyPurchOrder & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If declarations.MyRec.Fields("CC").Value > 0 Then
                        trycloseMyRec()
                        MyInvoiceUploaded = 1
                        Exit For
                    Else
                        trycloseMyRec()
                        MyInvoiceUploaded = 0
                    End If
                Next i
                If MyInvoiceUploaded = 0 Then
                    Main.button3.Enabled = True
                    Main.label6.Text = ""
                    Main.progressBar1.Minimum = 0
                    Main.progressBar1.Maximum = MyItemNodeList.Count - 1
                Else
                    Main.button3.Enabled = False
                    Main.label6.Text = "данная СФ уже загружена в Scala"
                End If
            End If

        Catch ex As Exception
            MsgBox("Ошибка " + ex.Message)
        End Try
    End Sub

    Private Sub RemoveCarrigeReturnSym(ByVal FileName As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление символов перевода каретки в файлах АББ
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim fs As New FileStream(FileName, FileMode.Open, FileAccess.ReadWrite)
        Dim fInfo As New FileInfo(FileName)
        Dim numBytes As Long = fInfo.Length
        Dim br As New BinaryReader(fs)
        Dim fileContent As Byte() = br.ReadBytes(CInt(numBytes))
        fs.SetLength(0)
        br.Close()

        Dim newContent As Byte() = New Byte(CInt(numBytes)) {}
        Dim j As Double = 0
        For i As Double = 0 To numBytes - 1
            If i < numBytes - 2 Then
                If fileContent(i) = &HD And fileContent(i + 1) = &HA Then
                    i = i + 1
                Else
                    newContent(j) = fileContent(i)
                    j = j + 1
                End If
            End If
        Next

        Dim fr As New FileStream(FileName, FileMode.Open, FileAccess.ReadWrite)
        fr.Write(newContent, 0, j)
        fr.Close()
    End Sub

    Public Sub UploadABBInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка файла с инвойсом ABB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String

        MyRezStr = ""

        LoadABBInvoiceToTMPTable()
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

    Private Sub LoadABBInvoiceToTMPTable()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка файла с инвойсом ABB во временную таблицу
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
        Dim MyInvoiceCurrExchRate As Double             '--Курс валюты в инвойсе
        Dim MyConsPurchaseOrderNum As String            '--Номер консолидированного заказа на закупку
        Dim MySupplierItemCode As String                '--код товара поставщика
        Dim MyQTY As Double                             '--количество
        Dim MySummWithoutVAT As Double                  '--Сумма без НДС за строку
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

        MyInvoice = MyHeaderNode.ChildNodes(0).InnerText
        MyInvoiceDate = Replace(MyHeaderNode.ChildNodes(1).InnerText, ".", "/")
        MyInvoiceCurrCode = CInt(Main.textBox5.Text)
        MySalesmanCode = declarations.SalesmanCode
        MySalesmanName = declarations.SalesmanName
        If aa.CurrentInfo.NumberDecimalSeparator = "," Then
            MyInvoiceCurrExchRate = CDbl(Replace(MyHeaderNode.ChildNodes(5).InnerText, ".", ","))
        Else
            MyInvoiceCurrExchRate = CDbl(MyHeaderNode.ChildNodes(5).InnerText)
        End If

        For i = 0 To MyItemNodeList.Count - 1
            'MyConsPurchaseOrderNum = Right("0000000000" + Right(MyItemNodeList(i).ChildNodes(13).InnerText, MyItemNodeList(i).ChildNodes(13).InnerText.Length - 3), 10)
            MyConsPurchaseOrderNum = Right("0000000000" + MyItemNodeList(i).ChildNodes(13).InnerText, 10)
            MySupplierItemCode = MyItemNodeList(i).ChildNodes(1).InnerText
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MyQTY = CDbl(Replace(MyItemNodeList(i).ChildNodes(3).InnerText, ".", ","))
            Else
                MyQTY = CDbl(MyItemNodeList(i).ChildNodes(3).InnerText)
            End If
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MySummWithoutVAT = CDbl(Replace(MyItemNodeList(i).ChildNodes(5).InnerText, ".", ","))
            Else
                MySummWithoutVAT = CDbl(MyItemNodeList(i).ChildNodes(5).InnerText)
            End If
            MyCountry = MyItemNodeList(i).ChildNodes(9).InnerText
            MyGTD = MyItemNodeList(i).ChildNodes(10).InnerText

            MySQLStr = "INSERT INTO #_MyInvoice "
            MySQLStr = MySQLStr & "(ID, Invoice, InvoiceDate, InvoiceCurrCode, SalesmanCode, SalesmanName, InvoiceCurrExchRate, "
            MySQLStr = MySQLStr & "ConsPurchaseOrderNum, SupplierItemCode, QTY, SummWithoutVAT, Country, GTD, RestQTY) "
            MySQLStr = MySQLStr & "VALUES (" & CStr(i + 1) & ", "
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

        Next i
    End Sub


End Module
