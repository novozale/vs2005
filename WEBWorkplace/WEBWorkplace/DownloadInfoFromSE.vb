Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.IO
Imports System.Drawing.Imaging

Public Class DownloadInfoFromSE

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Каталог для сохранения данных обязательно должен быть выбран.", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Код доступа к сервису обязательно должен быть введен.", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            TextBox3.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка картинок из выбранного каталога
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            DownloadInfo(Trim(TextBox1.Text), Trim(TextBox3.Text))
            MsgBox("Загрузка информации завершена.", MsgBoxStyle.Information, "Внимание!")
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор каталога с картинками
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub DownloadInfoFromSE_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// запрет выхода по alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub DownloadInfo(ByVal MyCatalog As String, ByVal MySecCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации с сервиса ШЕ
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer            '---счетчик
        
        If RadioButton1.Checked Then
            MySQLStr = "SELECT DISTINCT SC01060 AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01060 <> N'' AND SC01060 <> N'0') "
            MySQLStr = MySQLStr & "AND (SC01058 = N'0864' OR SC01058 = N'4974') "
            'MySQLStr = MySQLStr & "AND (SC01060 = N'A9MEM3110R')"
            MySQLStr = MySQLStr & "ORDER BY SC01060 "
        Else
            MySQLStr = "SELECT DISTINCT SC01060 AS CC "
            MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory "
            MySQLStr = MySQLStr & "WHERE (DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) "
            MySQLStr = MySQLStr & "AND ((PL01001 = N'0864') OR (PL01001 = N'4974')) "
            MySQLStr = MySQLStr & "ORDER BY SC01060 "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---Нет воообще ничего - ну и не будем ничего грузить
            trycloseMyRec()
            Exit Sub
        Else
            If My.Settings.UseOffice = "LibreOffice" Then
                Dim oServiceManager As Object
                Dim oDispatcher As Object
                Dim oDesktop As Object
                Dim oWorkBook As Object
                Dim oSheet As Object
                Dim oFrame As Object

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

                UploadHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
                    oSheet, oFrame)

                Declarations.MyRec.MoveLast()
                Label3.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    i = i + 1
                    DownloadOneItem(MyCatalog, MySecCode, Declarations.MyRec.Fields("CC").Value, i + 1, Nothing, _
                        oServiceManager, oDispatcher, oDesktop, oWorkBook, oSheet, oFrame)
                    Label2.Text = i
                    Application.DoEvents()
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()

                Dim args() As Object
                ReDim args(0)
                args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                args(0).Name = "ToPoint"
                args(0).Value = "$A$1"
                oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

                Dim args1() As Object
                ReDim args1(1)
                args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                args1(0).Name = "URL"
                args1(0).Value = "file:///" & Replace(MyCatalog, "\", "/") & "/ProductsFromScheiderElectric.ods"
                args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                args1(1).Name = "FilterName"
                args1(1).Value = "calc8"
                oDispatcher.executeDispatch(oFrame, ".uno:SaveAs", "", 0, args1)

                oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
                oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
            Else
                Dim MyObj As Object
                Dim MyWRKBook As Object

                MyObj = CreateObject("Excel.Application")
                MyObj.SheetsInNewWorkbook = 1
                MyWRKBook = MyObj.Workbooks.Add

                UploadHeaderExcel(MyWRKBook)

                Declarations.MyRec.MoveLast()
                Label3.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    i = i + 1
                    DownloadOneItem(MyCatalog, MySecCode, Declarations.MyRec.Fields("CC").Value, i + 1, MyWRKBook, _
                        Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    'Exit While
                    Label2.Text = i
                    Application.DoEvents()
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()

                MyObj.Application.DisplayAlerts = False
                MyWRKBook.SaveAs(MyCatalog & "\ProductsFromScheiderElectric.xlsx")
                MyObj.Application.DisplayAlerts = True

                MyObj.Application.Visible = True
                MyWRKBook = Nothing
                MyObj = Nothing
            End If
        End If
    End Sub

    Private Sub DownloadOneItem(ByVal MyCatalog As String, ByVal MySecCode As String, ByVal MyItemCode As String, ByVal MyCounter As Integer, ByRef MyWRKBook As Object, _
        ByRef oServiceManager As Object, ByRef oDispatcher As Object, ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации с сервиса ШЕ по 1 продукту
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyURL As String
        'Dim MySTR As String
        Dim webClient As New System.Net.WebClient
        Dim MyProxy As String

        '---прокси - если вдруг понадобится
        MyProxy = My.Settings.ProxyIP
        'MyProxy = "192.168.10.234:3128"
        Dim p As WebProxy = New WebProxy(MyProxy, True)
        'p.Credentials = New NetworkCredential("", "", "eskru")
        'p.Credentials = New NetworkCredential("novozhilov", "!564alexandr37", "eskru")
        p.Credentials = System.Net.CredentialCache.DefaultCredentials
        WebRequest.DefaultWebProxy = p

        webClient.Encoding = System.Text.Encoding.UTF8

        'MyURL = "http://web.se-ecatalog.ru/new-api/JSON/getdata?accessCode=" + MySecCode + "&commercialRef=" + MyItemCode
        'MyURL = "http://spbadm4:8880/web.se-ecatalog.ru/new-api/JSON/getdata?accessCode=" + MySecCode + "&commercialRef=" + MyItemCode
        MyURL = "http://spbadm4:8880/api.systeme.ru/new-api/JSON/getdata?accessCode=" + MySecCode + "&commercialRef=" + MyItemCode

        Try
            Dim result As String = webClient.DownloadString(MyURL)
            Dim json As JObject = JObject.Parse(result)

            '-----------------------картинки---------------------------------------------------
            Dim MySize As Double
            Dim MyType As String
            Dim MyPictureURL As String
            Dim CurrSize As Double

            MySize = 0
            MyType = ""
            MyPictureURL = ""

            Try
                Dim jarray As JArray = json.SelectToken("data").Item(0).SelectToken("images")
                For i As Integer = 0 To jarray.Count - 1
                    If json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("size") = "max" Then
                        CurrSize = 999999
                    Else
                        CurrSize = CDbl(json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("size"))
                    End If
                    If CurrSize > MySize Then
                        MySize = CurrSize
                        MyType = json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("type")
                        MyPictureURL = json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("url")
                    ElseIf CurrSize = MySize Then
                        If json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("type") = "JPG" And MyType <> "JPG" Then
                            MySize = CurrSize
                            MyType = json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("type")
                            MyPictureURL = json.SelectToken("data").Item(0).SelectToken("images").Item(i).SelectToken("url")
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            If MyPictureURL <> "" Then
                SavePicture(MyCatalog, MyItemCode, MyPictureURL)
            End If

            '-------------------название--------------------------------------------------------------
            Dim MyName As String

            MyName = ""

            Try
                MyName = json.SelectToken("data").Item(0).SelectToken("description")
                If MyName = Nothing Then
                    MyName = ""
                End If
            Catch ex As Exception
            End Try

            '----------------------описание--------------------------------------------------------------
            'etimru->features->feature...
            Dim MyDescr As String

            MyDescr = ""
            Try
                MyDescr = json.SelectToken("data").Item(0).SelectToken("eComeProductDescriptions")
                If MyDescr = Nothing Then
                    MyDescr = ""
                End If
            Catch ex As Exception
            End Try

            If Not MyWRKBook = Nothing Then
                WriteInfoToExcel(MyWRKBook, MyCounter, MyItemCode, MyName, MyDescr)
            Else
                WriteInfoToLO(oSheet, MyCounter, MyItemCode, MyName, MyDescr)
            End If

        Catch ex As Exception
            If Not MyWRKBook = Nothing Then
                WriteInfoToExcel(MyWRKBook, MyCounter, MyItemCode, "", "")
            Else
                WriteInfoToLO(oSheet, MyCounter, MyItemCode, "", "")
            End If
        End Try
    End Sub

    Private Sub SavePicture(ByVal MyCatalog As String, ByVal MyItemCode As String, ByVal MyPictureURL As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение одной картинки в каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Try
            Dim webClient As WebClient = New WebClient
            Dim data As Byte() = webClient.DownloadData(Replace(MyPictureURL, "https://", "http://spbadm4:8880/"))
            Dim fileExt = Path.GetExtension(MyPictureURL)

            Using mem = New MemoryStream(data)
                Using yourImage = Image.FromStream(mem)
                    If MyPictureURL.EndsWith(".jpg") Then
                        If File.Exists(MyCatalog + "\" + MyItemCode + ".jpg") Then File.Delete(MyCatalog + "\" + MyItemCode + ".jpg")
                        yourImage.Save(MyCatalog + "\" + MyItemCode + ".jpg", ImageFormat.Jpeg)
                    ElseIf MyPictureURL.EndsWith(".jpeg") Then
                        If File.Exists(MyCatalog + "\" + MyItemCode + ".jpeg") Then File.Delete(MyCatalog + "\" + MyItemCode + ".jpeg")
                        yourImage.Save(MyCatalog + "\" + MyItemCode + ".jpeg", ImageFormat.Jpeg)
                    ElseIf MyPictureURL.EndsWith(".gif") Then
                        If File.Exists(MyCatalog + "\" + MyItemCode + ".gif") Then File.Delete(MyCatalog + "\" + MyItemCode + ".gif")
                        yourImage.Save(MyCatalog + "\" + MyItemCode + ".gif", ImageFormat.Gif)
                    ElseIf MyPictureURL.EndsWith(".png") Then
                        If File.Exists(MyCatalog + "\" + MyItemCode + ".png") Then File.Delete(MyCatalog + "\" + MyItemCode + ".png")
                        yourImage.Save(MyCatalog + "\" + MyItemCode + ".png", ImageFormat.Png)
                    ElseIf MyPictureURL.EndsWith(".bmp") Then
                        If File.Exists(MyCatalog + "\" + MyItemCode + ".bmp") Then File.Delete(MyCatalog + "\" + MyItemCode + ".bmp")
                        yourImage.Save(MyCatalog + "\" + MyItemCode + ".bmp", ImageFormat.Bmp)
                    End If
                End Using
            End Using

        Catch ex As Exception
        End Try
    End Sub

    Private Sub UploadHeaderExcel(ByRef MyWRKBook As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка Excel файла
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A1") = "Код товара"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Range("B1") = "Название"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Range("C1") = "Описание"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 150

        MyWRKBook.ActiveSheet.Range("A1:C1").Select()
        MyWRKBook.ActiveSheet.Range("A1:C1").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A1:C1").Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A1:C1").WrapText = True
        With MyWRKBook.ActiveSheet.Range("A1:C1").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:C1").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:C1").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:C1").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:C1").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:C1").Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A1:C1").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A1:C1").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A1:C1").Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With
    End Sub

    Private Sub UploadHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка LibreOffice файла
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 7200
        oSheet.getColumns().getByName("C").Width = 28500

        i = 1
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код товара"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Название"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Описание"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
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

    End Sub

    Private Sub WriteInfoToExcel(ByRef MyWRKBook As Object, ByVal MyCounter As Integer, ByVal MyItemCode As String, ByVal MyName As String, ByVal MyDescr As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel информации по одному товару
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(MyCounter)).NumberFormat = "@"
        MyWRKBook.ActiveSheet.Range("A" & CStr(MyCounter)) = MyItemCode
        MyWRKBook.ActiveSheet.Range("B" & CStr(MyCounter)) = MyName
        MyWRKBook.ActiveSheet.Range("C" & CStr(MyCounter)) = MyDescr

    End Sub

    Private Sub WriteInfoToLO(ByRef oSheet As Object, ByVal MyCounter As Integer, ByVal MyItemCode As String, ByVal MyName As String, ByVal MyDescr As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в LibreOffice информации по одному товару
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        oSheet.getCellRangeByName("A" & CStr(MyCounter)).String = MyItemCode
        oSheet.getCellRangeByName("B" & CStr(MyCounter)).String = MyName
        oSheet.getCellRangeByName("C" & CStr(MyCounter)).String = MyDescr

    End Sub
End Class