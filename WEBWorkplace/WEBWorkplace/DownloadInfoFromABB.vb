Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.IO
Imports System.Drawing.Imaging

Public Class DownloadInfoFromABB

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

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

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Логин обязательно должен быть введен.", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            TextBox2.Select()
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

    Private Sub DownloadInfoFromABB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// запрет выхода по alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
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

    Private Sub DownloadInfo(ByVal MyCatalog As String, ByVal MySecCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации с сервиса ABB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyToken As String = ""
        Dim i As Integer            '---счетчик
        Dim MyObj As Object
        Dim MyWRKBook As Object

        MyToken = GetMyToken()
        If MyToken <> "" Then
            If RadioButton1.Checked Then
                MySQLStr = "SELECT DISTINCT SC01060 AS CC "
                MySQLStr = MySQLStr & "FROM SC010300 "
                MySQLStr = MySQLStr & "WHERE (SC01060 <> N'' AND SC01060 <> N'0') "
                MySQLStr = MySQLStr & "AND ((SC01058 = N'3046') OR (SC01058 = N'5832') OR (SC01058 = N'FIN103')) "
                MySQLStr = MySQLStr & "ORDER BY SC01060 "
            Else
                MySQLStr = "SELECT DISTINCT SC01060 AS CC "
                MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory "
                MySQLStr = MySQLStr & "WHERE (DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) "
                MySQLStr = MySQLStr & "AND ((PL01001 = N'3046') OR (PL01001 = N'5832') OR (PL01001 = N'FIN103'))"
                MySQLStr = MySQLStr & "ORDER BY SC01060 "
            End If

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '---Нет воообще ничего - ну и не будем ничего грузить
                trycloseMyRec()
                Exit Sub
            Else
                MyObj = CreateObject("Excel.Application")
                MyObj.SheetsInNewWorkbook = 1
                MyWRKBook = MyObj.Workbooks.Add

                UploadExcelHeader(MyWRKBook)

                Declarations.MyRec.MoveLast()
                Label3.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    i = i + 1
                    DownloadOneABBItem(MyCatalog, MyToken, Declarations.MyRec.Fields("CC").Value, i + 1, MyWRKBook)
                    Label2.Text = i
                    Application.DoEvents()
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()

                MyObj.Application.DisplayAlerts = False
                MyWRKBook.SaveAs(MyCatalog & "\ProductsFromABB.xlsx")
                MyObj.Application.DisplayAlerts = True

                MyObj.Application.Visible = True
                MyWRKBook = Nothing
                MyObj = Nothing
            End If
        Else
            MsgBox("невозможно получить token. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Function GetMyToken() As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение токена для работы с сайтом АББ  
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim request As System.Net.HttpWebRequest
        Dim writer As IO.BinaryWriter
        Dim response As System.Net.HttpWebResponse
        Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes("{""email"": """ + Trim(TextBox2.Text) + """,""password"": """ + Trim(TextBox3.Text) + """}")
        Dim responseContent As String
        Dim myToken As String
        Dim MyProxy As String

        Try
            '---прокси - если вдруг понадобится
            MyProxy = My.Settings.ProxyIP
            'MyProxy = "192.168.10.234:3128"
            Dim p As WebProxy = New WebProxy(MyProxy, True)
            'p.Credentials = New NetworkCredential("", "", "eskru")
            'p.Credentials = New NetworkCredential("novozhilov", "!564alexandr37", "eskru")
            p.Credentials = System.Net.CredentialCache.DefaultCredentials
            WebRequest.DefaultWebProxy = p

            'request = System.Net.WebRequest.Create("https://abb-api.ru/api/v1/users/authenticate")
            request = System.Net.WebRequest.Create("http://spbadm4:8880/abb-api.ru/api/v1/users/authenticate")
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = byteArray.Length

            writer = New IO.BinaryWriter(request.GetRequestStream())
            writer.Write(byteArray)
            writer.Close()

            response = request.GetResponse
            Dim reader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            responseContent = reader.ReadToEnd
            Dim json As JObject = JObject.Parse(responseContent)
            myToken = json.Item("token")
            GetMyToken = myToken
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Вниманеие!")
            GetMyToken = ""
        End Try
    End Function

    Private Sub DownloadOneABBItem(ByVal MyCatalog As String, ByVal MyToken As String, ByVal MyItemCode As String, _
        ByVal MyCounter As Integer, ByRef MyWRKBook As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации с сервиса ABB по 1 продукту
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim request As System.Net.HttpWebRequest
        Dim response As System.Net.HttpWebResponse
        Dim responseContent As String
        Dim MyPictureURL As String

        Try
            'request = System.Net.WebRequest.Create("http://abb-api.ru/api/v1/products/?manufacturerCode=" + MyItemCode)
            request = System.Net.WebRequest.Create("http://spbadm4:8880/abb-api.ru/api/v1/products?manufacturerCode=" + MyItemCode)
            request.Method = "GET"
            request.ContentType = "application/json"
            request.Headers.Add("authorization", "Bearer " + MyToken)

            response = request.GetResponse
            Dim reader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            responseContent = reader.ReadToEnd
            Dim json As JObject = JObject.Parse(responseContent)

            '--------------------------картинки------------------------------------
            MyPictureURL = ""
            Dim jarray As JArray = json.SelectToken("items").Item(0).SelectToken("images")
            For i As Integer = 0 To jarray.Count - 1
                If json.SelectToken("items").Item(0).SelectToken("images").Item(i).SelectToken("size") = "max" Then
                    MyPictureURL = json.SelectToken("items").Item(0).SelectToken("images").Item(i).SelectToken("url")
                End If
            Next
            If MyPictureURL <> "" Then
                SavePicture(MyCatalog, MyItemCode, MyPictureURL)
            End If

            '-------------------название--------------------------------------------------------------
            Dim MyName As String

            MyName = ""

            Try
                MyName = json.SelectToken("items").Item(0).SelectToken("name").Item(0).SelectToken("value")
                If MyName = Nothing Then
                    MyName = ""
                End If
            Catch ex As Exception
            End Try

            WriteInfoToExcel(MyWRKBook, MyCounter, MyItemCode, MyName, "")
        Catch ex As Exception
            WriteInfoToExcel(MyWRKBook, MyCounter, MyItemCode, "", "")
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

    Private Sub UploadExcelHeader(ByRef MyWRKBook As Object)
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
End Class