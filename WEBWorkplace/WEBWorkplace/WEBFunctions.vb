Imports System
Imports System.IO
Imports System.Net



Module WEBFunctions

    Public Function GetFolderPath() As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MainForm.FolderBrowserDialog1.ShowDialog()
        GetFolderPath = MainForm.FolderBrowserDialog1.SelectedPath
    End Function


    Public Sub UploadToFile(ByVal MyFilename As String, ByVal MyCatalog As String, ByVal MyFullUploadFlag As Integer, ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � �������� � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyFilename = "Salesmans.txt" Then                        '----------��������
            MySQLStr = "exec spp_WEB_Salesmans_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "ItemGroups.txt" Then                   '----------������ �������
            MySQLStr = "exec spp_WEB_ItemGroups_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "ItemSubGroups.txt" Then                '----------��������� �������
            MySQLStr = "exec spp_WEB_ItemSubGroups_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "Items.txt" Then                        '----------������
            MySQLStr = "exec spp_WEB_Items_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "Manufacturers.txt" Then                '----------�������������
            MySQLStr = "exec spp_WEB_Manufacturers_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "CardDiscounts.txt" Then                '----------����� ������ ���������, ���������� ����� WEB ����
            MySQLStr = "exec spp_WEB_CardDiscounts_FromDB "
        ElseIf MyFilename = "Price.txt" Then                        '----------����� �����
            MySQLStr = "exec spp_WEB_Price_FromDB "
        ElseIf MyFilename = "ShopsAvailability.txt" Then            '----------����������� �� �������
            MySQLStr = "exec spp_WEB_ShopAvailability_FromDB "
        ElseIf MyFilename = "CardSales.txt" Then                    '----------������� ���������
            MySQLStr = "exec spp_WEB_CardSales_FromDB " & CStr(MyFullUploadFlag) & ", " & CStr(MarkFlag)
        ElseIf MyFilename = "CardSalesDetails.txt" Then             '----------������� ������
            MySQLStr = "exec spp_WEB_CardSalesDetails_FromDB " & CStr(MyFullUploadFlag)
        ElseIf MyFilename = "FullUpload.txt" Then                   '----------���� - ���� "������ ��������"
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
        '// �������� ���������� � ����
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
        '// ������ �������� �� �� � ��������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
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
        '// "������" �������� �� �� � ��������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
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
        '// �������� ���������� � ����������� �� ������� �� �� ������ � ������������ �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
        Else
            UploadToFile("ShopsAvailability.txt", MyCatalog, 0, MarkFlag)
        End If
    End Sub

    Public Sub SalesUploadToCatalog(ByVal MarkFlag As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � �������� �� �� � ��������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLstr As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
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
        '// ������ �������� �� �� � ��������� ������� � ��������� �� WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        FullUploadToCatalog_WEB = False

        SetDBBlock()        '--------------���������� �������� ��� ��������� �������������
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
        RemoveDBBlock()     '--------------������ ���������� �������� ��� ��������� �������������
    End Function

    Public Function DailyUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ (����������) �������� �� �� ������ � ������������ ������� � ��������� �� WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        DailyUploadToCatalog_WEB = False

        SetDBBlock()        '--------------���������� �������� ��� ��������� �������������
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
        RemoveDBBlock()     '--------------������ ���������� �������� ��� ��������� �������������
    End Function

    Public Function SalesUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� �� �� � ��������� ������� � ��������� �� WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MySQLStr As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        SalesUploadToCatalog_WEB = False

        SetDBBlock()        '--------------���������� �������� ��� ��������� �������������
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
        RemoveDBBlock()     '--------------������ ���������� �������� ��� ��������� �������������
    End Function

    Public Function AvailabilityUploadToCatalog_WEB(ByVal MarkFlag As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� �� ������� �� �� � ��������� ������� � ��������� �� WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String
        Dim MyArchive As String

        MyCatalog = My.Settings.ExchangeCatalog
        AvailabilityUploadToCatalog_WEB = False

        SetDBBlock()        '--------------���������� �������� ��� ��������� �������������
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
        RemoveDBBlock()     '--------------������ ���������� �������� ��� ��������� �������������
    End Function

    Public Function PrepareCatalogForExchange(ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� (�������� � �������) �������� ��� �������� �� WEB
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim AllFilesMask As String
        Dim ArchivatorFile As String

        If Directory.Exists(MyCatalog) = False Then
            '-----�������� ��������
            Try
                Dim di As DirectoryInfo = Directory.CreateDirectory(MyCatalog)
            Catch ex As Exception
                MsgBox("�� ���� ������� ������� " & MyCatalog & " ��� �������� ������ �� WEB. " & ex.Message, MsgBoxStyle.Critical, "��������!")
                PrepareCatalogForExchange = False
                Exit Function
            End Try
        End If

        '-----������� ��������
        Try
            AllFilesMask = MyCatalog & "\*.txt"
            If Len(Dir(AllFilesMask)) > 0 Then '���� ���� �� ���� ����
                Kill(AllFilesMask)
            End If
        Catch ex As Exception
            MsgBox("�� ���� �������� ������� " & MyCatalog & " ��� �������� ������ �� WEB. " & ex.Message, MsgBoxStyle.Critical, "��������!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        Try
            AllFilesMask = MyCatalog & "\*.rar"
            If Len(Dir(AllFilesMask)) > 0 Then '���� ���� �� ���� ����
                Kill(AllFilesMask)
            End If
        Catch ex As Exception
            MsgBox("�� ���� �������� ������� " & MyCatalog & " ��� �������� ������ �� WEB. " & ex.Message, MsgBoxStyle.Critical, "��������!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        '-----�������� ������� ����������
        Try
            ArchivatorFile = My.Settings.ArchivatorPath & "\Rar.exe"
            Dim FI As IO.FileInfo = New IO.FileInfo(ArchivatorFile)
            If FI.Exists Then
            Else
                MsgBox("�� ���� ����� ��������� Rar.exe � �������� " & My.Settings.ArchivatorPath & ". ", MsgBoxStyle.Critical, "��������!")
                PrepareCatalogForExchange = False
                Exit Function
            End If
        Catch ex As Exception
            MsgBox("�� ���� ����� ��������� Rar.exe � �������� " & My.Settings.ArchivatorPath & ". " & ex.Message, MsgBoxStyle.Critical, "��������!")
            PrepareCatalogForExchange = False
            Exit Function
        End Try

        PrepareCatalogForExchange = True
    End Function

    Public Function PrepareArchive(ByVal ArchiveName As String, ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������������ ������ �� ������ *.txt � �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim FI As IO.FileInfo = New IO.FileInfo(MyCatalog & "\" & ArchiveName)
        If FI.Exists Then
            Try
                FI.Delete()
            Catch ex As Exception
                MsgBox("�� ���� ������� ���� " & FI.Name & ". " & ex.Message, MsgBoxStyle.Critical, "��������!")
                PrepareArchive = False
                Exit Function
            End Try
        End If

        Try
            Shell("""" & My.Settings.ArchivatorPath & "\Rar.exe""" & " m5 -ep " & """" & MyCatalog & "\" & ArchiveName & """ """ & MyCatalog & "\*.txt""", AppWinStyle.Hide, True)
        Catch ex As Exception
            MsgBox("������ �������� ������ " & ArchiveName & ". " & ex.Message, MsgBoxStyle.Critical, "��������!")
            PrepareArchive = False
            Exit Function
        End Try

        PrepareArchive = True
    End Function

    Public Function SendToFTP(ByVal ArchiveName As String, ByVal MyCatalog As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� �� FTP
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim ServerURL As String
        Dim FTPLogin As String
        Dim FTPPassword As String

        ServerURL = My.Settings.ServerURL
        FTPLogin = My.Settings.FTPLogin
        FTPPassword = My.Settings.FTPPassword

        Try
            '-----��������� �����
            Dim request As WebClient = New WebClient()
            request.Proxy = Nothing
            request.Credentials = New NetworkCredential(FTPLogin, FTPPassword)
            request.UploadFile(ServerURL & "/" & ArchiveName, MyCatalog & "\" & ArchiveName)
            '-----��������� ���� � ������ ������ + .txt ������ ������ "����"
            Dim f As New StreamWriter(MyCatalog & "\" & ArchiveName & ".txt", False, System.Text.Encoding.GetEncoding(1251))
            f.Write("Flag")
            f.Close()
            request.UploadFile(ServerURL & "/" & ArchiveName & ".txt", MyCatalog & "\" & ArchiveName & ".txt")
        Catch ex As WebException
            MsgBox("������ �������� �� FTP " & ex.Message, MsgBoxStyle.Critical, "��������!")
            SendToFTP = False
            Exit Function
        End Try

        SendToFTP = True
    End Function

    Public Function GetFTPConfirmation(ByVal ArchiveName As String, ByVal UploadType As Integer) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������������� �������� �� FTP ��� ��������� ������
        '// UploadType  4 ������ ��������
        '//             2 ������ ��������
        '//             3 ���������� � ��������
        '//             1 ��������� �������
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
            MsgBox("������ ��������� ������������� � �������� �� FTP. ", MsgBoxStyle.Critical, "��������!")
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
        '// ���������� ������ � �� � ������ ������������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If UploadType = 4 Then          '-------------������ ��������
            '-----------��������
            MySQLStr = "UPDATE tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus, ScalaStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������ �������
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------��������� �������
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------�������������
            MySQLStr = "UPDATE tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������� ������
            MySQLStr = "DELETE FROM tbl_WEB_CardSalesDetails "
            MySQLStr = MySQLStr & "FROM tbl_WEB_CardSalesDetails INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_CardSales ON tbl_WEB_CardSalesDetails.ClientCode = tbl_WEB_CardSales.ClientCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_CardSales.WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������� ���������
            MySQLStr = "UPDATE tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������������� + ������ (���������� �������� ��� ������������ ��������� �� ������ � �������������, �� � � ������)
            MySQLStr = "truncate table dbo.tbl_WEB_ManufacturersPlusCountries "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO dbo.tbl_WEB_ManufacturersPlusCountries "
            MySQLStr = MySQLStr & "SELECT DISTINCT ManufacturerCode, CountryCode "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (SubGroupCode <> N'') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        ElseIf UploadType = 3 Then      '-------------�������� ���������� � ��������
            '----------������� ������
            MySQLStr = "DELETE FROM tbl_WEB_CardSalesDetails "
            MySQLStr = MySQLStr & "FROM tbl_WEB_CardSalesDetails INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_CardSales ON tbl_WEB_CardSalesDetails.ClientCode = tbl_WEB_CardSales.ClientCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_CardSales.WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������� ���������
            MySQLStr = "UPDATE tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_CardSales "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        ElseIf UploadType = 2 Then      '-------------������ �������� 
            '-----------��������
            MySQLStr = "UPDATE tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus, ScalaStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������ �������
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------��������� �������
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������
            MySQLStr = "UPDATE tbl_WEB_Items "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Items "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------�������������
            MySQLStr = "UPDATE tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "SET WEBStatus = RMStatus "
            MySQLStr = MySQLStr & "WHERE (WEBStatus <> 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (WEBStatus = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----------������������� + ������ (���������� �������� ��� ������������ ��������� �� ������ � �������������, �� � � ������)
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
        '// ��������� ������ ���������� �� �������� ������
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
        '// ������ ������ ���������� ����� �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_WEB_RemoveBlock "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Module
