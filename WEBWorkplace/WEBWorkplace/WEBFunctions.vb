Imports System
Imports System.IO
Imports System.Net



Module WEBFunctions

    Public Function GetFolderPath() As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выбор каталога
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MainForm.FolderBrowserDialog1.ShowDialog()
        GetFolderPath = MainForm.FolderBrowserDialog1.SelectedPath
    End Function


    Public Sub UploadToFile(ByVal MyFilename As String, ByVal MyCatalog As String, ByVal MyFullUploadFlag As Integer, ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// подготовка рекордсета к выгрузке в файл
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyFilename = "Salesmans.txt" Then                        '----------продавцы
            MySQLStr = "exec spp_WEB_Salesmans_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "ItemGroups.txt" Then                   '----------группы товаров
            MySQLStr = "exec spp_WEB_ItemGroups_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "ItemSubGroups.txt" Then                '----------подгруппы товаров
            MySQLStr = "exec spp_WEB_ItemSubGroups_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "Items.txt" Then                        '----------товары
            MySQLStr = "exec spp_WEB_Items_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "Manufacturers.txt" Then                '----------производители
            MySQLStr = "exec spp_WEB_Manufacturers_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "CardDiscounts.txt" Then                '----------Общие скидки клииентам, работающим через WEB сайт
            MySQLStr = "exec spp_WEB_CardDiscounts_FromDB "
        ElseIf MyFilename = "Price.txt" Then                        '----------Прайс листы
            MySQLStr = "exec spp_WEB_Price_FromDB "
        ElseIf MyFilename = "ShopsAvailability.txt" Then            '----------Доступность на складах
            MySQLStr = "exec spp_WEB_ShopAvailability_FromDB "
        ElseIf MyFilename = "CardSales.txt" Then                    '----------Продажи заголовок
            MySQLStr = "exec spp_WEB_CardSales_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "CardSalesDetails.txt" Then             '----------продажи строки
            MySQLStr = "exec spp_WEB_CardSalesDetails_FromDB " & CStr(MyFullUploadFlag)
        ElseIf MyFilename = "FullUpload.txt" Then                   '----------Файл - флаг "полная выгрузка"
            MySQLStr = "SELECT '""0""' AS CC "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            ExportToCSV(MyFilename, MyCatalog, Declarations.MyRec)
        End If
        trycloseMyRec()
    End Sub

    Public Sub ExportToCSV(ByVal MyFilename As String, ByVal MyCatalog As String, ByVal MyRecordset As ADODB.Recordset)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка рекордсета в файл
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyStr As String

        Dim f As New StreamWriter(MyCatalog & "\" & MyFilename, False, System.Text.Encoding.GetEncoding(1251))
        MyStr = MyRecordset.GetString(, , ";", vbCrLf)
        f.Write(MyStr)
        f.Close()

    End Sub

    Public Sub FullUploadToCatalog(ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка из БД в выбранный каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            'MySQLStr = "exec spp_WEB_Sales_FromScala "
            MySQLStr = "exec spp_WEB_ALL_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            UploadToFile("ItemGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("ItemSubGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Items.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Salesmans.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Manufacturers.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardDiscounts.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Price.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("ShopsAvailability.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardSales.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardSalesDetails.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("FullUpload.txt", MyCatalog, 1, MarkFlag)
        End If
    End Sub

    Public Sub NightUploadToCatalog(ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// "Ночная" выгрузка из БД в выбранный каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            MySQLStr = "exec spp_WEB_ALL_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            UploadToFile("ItemGroups.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("ItemSubGroups.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("Items.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("Salesmans.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("Manufacturers.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("CardDiscounts.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("Price.txt", MyCatalog, 0, MarkFlag)
        End If
    End Sub

    Public Sub AvailabilityUploadToCatalog(ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации о доступности на складах из БД файлов в определенный каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            UploadToFile("ShopsAvailability.txt", MyCatalog, 0, MarkFlag)
        End If
    End Sub

    Public Sub SalesUploadToCatalog(ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выгрузка информации о продажах из БД в выбранный каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLstr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            MySQLstr = "exec spp_WEB_Sales_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLstr)

            UploadToFile("CardSales.txt", MyCatalog, 0, MarkFlag)
            UploadToFile("CardSalesDetails.txt", MyCatalog, 0, MarkFlag)
        End If
    End Sub

    Public Function FullUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка из БД в выбранный каталог с отправкой на WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        FullUploadToCatalog_WEB = False

        SetDBBlock()        '--------------Блокировка выгрузки для остальных пользователей
        If PrepareCatalogForExchange(MyCatalog) = True Then
            'MySQLStr = "exec spp_WEB_Sales_FromScala "
            MySQLStr = "exec spp_WEB_ALL_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            UploadToFile("ItemGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("ItemSubGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Items.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Salesmans.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Manufacturers.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardDiscounts.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Price.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("ShopsAvailability.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardSales.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardSalesDetails.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("FullUpload.txt", MyCatalog, 1, MarkFlag)

            MyArchive = "FullUpload_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".rar"

            If PrepareArchive(MyArchive, MyCatalog) = True Then
                If SendToFTP(MyArchive, MyCatalog) = True Then
                    If GetFTPConfirmation(MyArchive, 4) = True Then
                        FullUploadToCatalog_WEB = True
                    End If
                End If
            End If
        End If
        RemoveDBBlock()     '--------------Снятие блокировки выгрузки для остальных пользователей
    End Function

    Public Function DailyUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ночная (ежедневная) выгрузка из БД файлов в определенный каталог с отправкой на WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        DailyUploadToCatalog_WEB = False

        SetDBBlock()        '--------------Блокировка выгрузки для остальных пользователей
        If PrepareCatalogForExchange(MyCatalog) = True Then
            'MySQLStr = "exec spp_WEB_Sales_FromScala "
            MySQLStr = "exec spp_WEB_ALL_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            UploadToFile("ItemGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("ItemSubGroups.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Items.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Salesmans.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Manufacturers.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardDiscounts.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("Price.txt", MyCatalog, 1, MarkFlag)

            MyArchive = "FullUpload_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".rar"

            If PrepareArchive(MyArchive, MyCatalog) = True Then
                If SendToFTP(MyArchive, MyCatalog) = True Then
                    If GetFTPConfirmation(MyArchive, 2) = True Then
                        DailyUploadToCatalog_WEB = True
                    End If
                End If
            End If
        End If
        RemoveDBBlock()     '--------------Снятие блокировки выгрузки для остальных пользователей
    End Function

    Public Function SalesUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка из БД в выбранный каталог с отправкой на WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        SalesUploadToCatalog_WEB = False

        SetDBBlock()        '--------------Блокировка выгрузки для остальных пользователей
        If PrepareCatalogForExchange(MyCatalog) = True Then
            MySQLStr = "exec spp_WEB_Sales_FromScala "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            UploadToFile("CardSales.txt", MyCatalog, 1, MarkFlag)
            UploadToFile("CardSalesDetails.txt", MyCatalog, 1, MarkFlag)

            MyArchive = "PartUpload_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".rar"

            If PrepareArchive(MyArchive, MyCatalog) = True Then
                If SendToFTP(MyArchive, MyCatalog) = True Then
                    If GetFTPConfirmation(MyArchive, 3) = True Then
                        SalesUploadToCatalog_WEB = True
                    End If
                End If
            End If
        End If
        RemoveDBBlock()     '--------------Снятие блокировки выгрузки для остальных пользователей
    End Function

    Public Function AvailabilityUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка доступности на складах из БД в выбранный каталог с отправкой на WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        AvailabilityUploadToCatalog_WEB = False

        SetDBBlock()        '--------------Блокировка выгрузки для остальных пользователей
        If PrepareCatalogForExchange(MyCatalog) = True Then
            UploadToFile("ShopsAvailability.txt", MyCatalog, 1, MarkFlag)

            MyArchive = "PartUpload_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".rar"

            If PrepareArchive(MyArchive, MyCatalog) = True Then
                If SendToFTP(MyArchive, MyCatalog) = True Then
                    If GetFTPConfirmation(MyArchive, 1) = True Then
                        AvailabilityUploadToCatalog_WEB = True
                    End If
                End If
            End If
        End If
        RemoveDBBlock()     '--------------Снятие блокировки выгрузки для остальных пользователей
    End Function

    Public Function PrepareCatalogForExchange(ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подготовка (создание и очистка) каталога для выгрузок на WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim AllFilesMask As String
        Dim ArchivatorFile As String

        If Directory.Exists(MyCatalog) = False Then
            '-----создание каталога
            Try
                Dim di As DirectoryInfo = Directory.CreateDirectory(MyCatalog)
            Catch ex As Exception
                MsgBox("не могу создать каталог " & MyCatalog & " для выгрузки данных на WEB. " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                PrepareCatalogForExchange = False
                Exit Function
            End Try
        End If

        '-----очистка каталога
        Try
            AllFilesMask = MyCatalog & "\*.txt"
            If Len(Dir(AllFilesMask)) > 0 Then 'есть хотя бы один файл
                Kill(AllFilesMask)
            End If
        Catch ex As Exception
            MsgBox("не могу очистить каталог " & MyCatalog & " для выгрузки данных на WEB. " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        Try
            AllFilesMask = MyCatalog & "\*.rar"
            If Len(Dir(AllFilesMask)) > 0 Then 'есть хотя бы один файл
                Kill(AllFilesMask)
            End If
        Catch ex As Exception
            MsgBox("не могу очистить каталог " & MyCatalog & " для выгрузки данных на WEB. " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        '-----проверка наличия архиватора
        Try
            ArchivatorFile = My.Settings.ArchivatorPath & "\Rar.exe"
            Dim FI As IO.FileInfo = New IO.FileInfo(ArchivatorFile)
            If FI.Exists Then
            Else
                MsgBox("не могу найти архиватор Rar.exe в каталоге " & My.Settings.ArchivatorPath & ". ", MsgBoxStyle.Critical, "Внимание!")
                PrepareCatalogForExchange = False
                Exit Function
            End If
        Catch ex As Exception
            MsgBox("не могу найти архиватор Rar.exe в каталоге " & My.Settings.ArchivatorPath & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        PrepareCatalogForExchange = True
    End Function

    Public Function PrepareArchive(ByVal ArchiveName As String, ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование архива из файлов *.txt в заданном каталоге
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim FI As IO.FileInfo = New IO.FileInfo(MyCatalog & "\" & ArchiveName)
        If FI.Exists Then
            Try
                FI.Delete()
            Catch ex As Exception
                MsgBox("не могу удалить файл " & FI.Name & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                PrepareArchive = False
                Exit Function
            End Try
        End If

        Try
            Shell("""" & My.Settings.ArchivatorPath & "\Rar.exe""" & " m5 -ep " & """" & MyCatalog & "\" & ArchiveName & """ """ & MyCatalog & "\*.txt""", AppWinStyle.Hide, True)
        Catch ex As Exception
            MsgBox("ошибка создания архива " & ArchiveName & ". " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            PrepareArchive = False
            Exit Function
        End Try

        PrepareArchive = True
    End Function

    Public Function SendToFTP(ByVal ArchiveName As String, ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка файла на FTP
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim ServerURL As String
        Dim FTPLogin As String
        Dim FTPPassword As String

        ServerURL = My.Settings.ServerURL
        FTPLogin = My.Settings.FTPLogin
        FTPPassword = My.Settings.FTPPassword

        Try
            '-----выгружаем архив
            Dim request As WebClient = New WebClient()
            request.Proxy = Nothing
            request.Credentials = New NetworkCredential(FTPLogin, FTPPassword)
            request.UploadFile(ServerURL & "/" & ArchiveName, MyCatalog & "\" & ArchiveName)
            '-----выгружаем файл с именем архива + .txt внутри запись "флаг"
            Dim f As New StreamWriter(MyCatalog & "\" & ArchiveName & ".txt", False, System.Text.Encoding.GetEncoding(1251))
            f.Write("Flag")
            f.Close()
            request.UploadFile(ServerURL & "/" & ArchiveName & ".txt", MyCatalog & "\" & ArchiveName & ".txt")
        Catch ex As WebException
            MsgBox("Ошибка выгрузки на FTP " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            SendToFTP = False
            Exit Function
        End Try

        SendToFTP = True
    End Function

    Public Function GetFTPConfirmation(ByVal ArchiveName As String, ByVal UploadType As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение подтверждения загрузки на FTP или получение ошибки
        '// UploadType  4 полная загрузка
        '//             2 ночная выгрузка
        '//             3 информация о продажах
        '//             1 состояние складов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim ServerURL As String
        Dim FTPLogin As String
        Dim FTPPassword As String
        Dim ErrFlag As Integer
        Dim i As Integer
        Dim j As Integer
        Dim MyDelay As Integer

        ServerURL = My.Settings.ServerURL
        FTPLogin = My.Settings.FTPLogin
        FTPPassword = My.Settings.FTPPassword
        j = My.Settings.NumberTry
        MyDelay = My.Settings.DelayTime

        For i = 0 To j
            Try
                Dim request As FtpWebRequest = DirectCast(WebRequest.Create(ServerURL & "/" & ArchiveName & ".ok"), FtpWebRequest)
                request.Proxy = Nothing
                request.Credentials = New NetworkCredential(FTPLogin, FTPPassword)
                request.Method = WebRequestMethods.Ftp.DeleteFile
                Dim ftpResp As FtpWebResponse = request.GetResponse
                ErrFlag = 0
                Exit For
            Catch ex As Exception
                ErrFlag = 1
            End Try

            Try
                Dim request As FtpWebRequest = DirectCast(WebRequest.Create(ServerURL & "/" & ArchiveName & ".err"), FtpWebRequest)
                request.Proxy = Nothing
                request.Credentials = New NetworkCredential(FTPLogin, FTPPassword)
                request.Method = WebRequestMethods.Ftp.DeleteFile
                Dim ftpResp As FtpWebResponse = request.GetResponse
                ErrFlag = 0
                Exit For
            Catch ex As Exception
                ErrFlag = 1
            End Try

            Application.DoEvents()
            Threading.Thread.Sleep(MyDelay)
        Next i

        If ErrFlag = 1 Then
            MsgBox("Ошибка получения подтверждения о выгрузке на FTP. ", MsgBoxStyle.Critical, "Внимание!")
            GetFTPConfirmation = False
            Exit Function
        Else
            UpdateConfirmedInfo(UploadType)
        End If

        GetFTPConfirmation = True
    End Function

    Public Sub UpdateConfirmedInfo(ByVal UploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление флагов в БД в случае положительной загрузки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If UploadType = 4 Then          '-------------полная выгрузка
            '-----------продавцы
            MySQLStr = "UPDATE tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus, ScalaStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------группы товаров
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------подгруппы товаров
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------товары
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------производители
            MySQLStr = "UPDATE tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------продажи строки
            MySQLStr = "DELETE FROM tbl_WEB_CardSalesDetails "
            MySQLStr = MySQLStr & "FROM tbl_WEB_CardSalesDetails INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_CardSales ON tbl_WEB_CardSalesDetails.ClientCode = tbl_WEB_CardSales.ClientCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_CardSales.WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------Продажи заголовок
            MySQLStr = "UPDATE tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------Производители + страны (обновление таблички для отслеживания изменений не только в производителе, но и в связке)
            MySQLStr = "truncate table dbo.tbl_WEB_ManufacturersPlusCountries "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO dbo.tbl_WEB_ManufacturersPlusCountries "
            MySQLStr = MySQLStr & "SELECT DISTINCT ManufacturerCode, CountryCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (SubGroupCode <> N'') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        ElseIf UploadType = 3 Then      '-------------выгрузка информации о продажах
            '----------продажи строки
            MySQLStr = "DELETE FROM tbl_WEB_CardSalesDetails "
            MySQLStr = MySQLStr & "FROM tbl_WEB_CardSalesDetails INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_CardSales ON tbl_WEB_CardSalesDetails.ClientCode = tbl_WEB_CardSales.ClientCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_CardSales.WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------Продажи заголовок
            MySQLStr = "UPDATE tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        ElseIf UploadType = 2 Then      '-------------ночная выгрузка 
            '-----------продавцы
            MySQLStr = "UPDATE tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus, ScalaStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------группы товаров
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------подгруппы товаров
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------товары
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------производители
            MySQLStr = "UPDATE tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------Производители + страны (обновление таблички для отслеживания изменений не только в производителе, но и в связке)
            MySQLStr = "truncate table dbo.tbl_WEB_ManufacturersPlusCountries "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO dbo.tbl_WEB_ManufacturersPlusCountries "
            MySQLStr = MySQLStr & "SELECT DISTINCT ManufacturerCode, CountryCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (SubGroupCode <> N'') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
    End Sub

    Public Sub SetDBBlock()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка записи блокировки до выгрузки данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_WEB_SetBlock N'" & Declarations.UserCode & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub


    Public Sub RemoveDBBlock()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Снятие записи блокировки после выгрузки данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_WEB_RemoveBlock "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Module
