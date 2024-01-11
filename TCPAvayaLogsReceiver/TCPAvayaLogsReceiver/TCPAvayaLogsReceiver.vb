Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.IO
Imports System.Net.Mail

Public Class TCPAvayaLogsReceiver
    Public Listener As TcpListener
    Private LIP As String = "127.0.0.1"
    Private LP As String = "50000"
    Private MySocketListenerThread As Threading.Thread
    Private MyFileLookedThread As Threading.Thread
    Private DBC As MyDBConnector
    Structure DBPutParam
        Public IPAddr As String
        Public ArrEl() As String
    End Structure
    Private MyMSGRests As New Dictionary(Of String, String)

    Protected Overrides Sub OnStart(ByVal args() As String)
        GetLocalIP()
        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "OnStart 1 ->Service IP is " & LIP)
        End If
        GetLocalP()
        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "OnStart 2 ->Service Port is " & LP)
        End If

        Try
            DBC = New MyDBConnector
            If DBC.InitMyConn() = False Then
                '---уведомление о невозможности установления соединения с БД
                SendMyReminder("ошибка запуска сервиса TCPAvayaLogsReceiver", "Невозможно установить сообщение с БД. Подробности в Лог файле" & vbCrLf)
                Me.Stop()
            End If
            DBC.CloseMyConn()
            DBC = Nothing
        Catch ex As Exception
            '---уведомление о невозможности создания БД коннектора
            SendMyReminder("ошибка запуска сервиса TCPAvayaLogsReceiver", "Сообщение об ошибке:" & vbCrLf & ex.Message)
            EventLog.WriteEntry("TCPAvayaLogsReceiver", ex.Message)
            Me.Stop()
        End Try
        Try
            Listener = New TcpListener(IPAddress.Parse(LIP), LP)
            Listener.Start()
        Catch ex As Exception
            '---уведомление о невозможности запуска сервиса
            SendMyReminder("ошибка запуска сервиса TCPAvayaLogsReceiver", "Сообщение об ошибке:" & vbCrLf & ex.Message)
            EventLog.WriteEntry("TCPAvayaLogsReceiver", ex.Message)
            Me.Stop()
        End Try
        Try
            MySocketListenerThread = New Threading.Thread(AddressOf MySocketListener)
            MySocketListenerThread.IsBackground = True
            MySocketListenerThread.Start()
        Catch ex As Exception
            '---уведомление о невозможности запуска сервиса
            SendMyReminder("ошибка запуска сервиса TCPAvayaLogsReceiver", "Сообщение об ошибке:" & vbCrLf & ex.Message)
            EventLog.WriteEntry("TCPAvayaLogsReceiver", ex.Message)
            Me.Stop()
        End Try

        Try
            MyFileLookedThread = New Threading.Thread(AddressOf PutLogFilesToDB)
            MyFileLookedThread.IsBackground = True
            MyFileLookedThread.Start()
        Catch ex As Exception
            '---уведомление о невозможности запуска сервиса
            SendMyReminder("ошибка запуска сервиса TCPAvayaLogsReceiver", "Сообщение об ошибке:" & vbCrLf & ex.Message)
            EventLog.WriteEntry("TCPAvayaLogsReceiver", ex.Message)
            Me.Stop()
        End Try

        '-----надо если поток для записи данных в БД не запускается
        'DBC = New MyDBConnector

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.

    End Sub

    Private Sub GetLocalIP()
        Try
            LIP = My.Settings.TCPAddress
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "GetLocalIP 1 ->" & ex.Message)
        End Try
    End Sub

    Private Sub GetLocalP()
        Try
            LP = My.Settings.TCPPort
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "GetLocalIP 2 ->" & ex.Message)
        End Try
    End Sub

    Public Sub MySocketListener()
        Dim IncomingClient As TcpClient
        Dim IncomingIP As String
        Dim RemoteEndPointParts As String()
        Try
            While (True)
                IncomingClient = Listener.AcceptTcpClient
                IncomingIP = IncomingClient.Client.RemoteEndPoint.ToString
                RemoteEndPointParts = IncomingIP.Split(":")
                IncomingIP = RemoteEndPointParts(0)
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("TCPAvayaLogsReceiver", "MySocketListener 1 -> Request from IP: " & IncomingIP)
                End If
                Dim connClient As New ConnectedClient(IncomingClient, IncomingIP)
                AddHandler connClient.dataReceived, AddressOf Me.messageReceived
                AddHandler connClient.ClientDisconnected, AddressOf Me.ClientDisconnected
                If MyMSGRests.ContainsKey(IncomingIP) Then
                Else
                    MyMSGRests.Add(IncomingIP, "")
                End If
            End While
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "MySocketListener 2 -> " & ex.Message)
            Me.Stop()
        End Try
    End Sub

    Private Sub messageReceived(ByVal sender As ConnectedClient, ByVal message As String)
        Dim StrArray() As String
        Dim SubStrArray() As String
        Dim i As Integer
        Dim MyStr As String
        Dim MyParam As DBPutParam

        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "messageReceived 1 -> " & "messageReceived от IP " & sender.IncomingIP & " >>" & message)
        End If

        '---Дополнительно - проверяем, нет ли остатка от предыдущего сообщения
        If MyMSGRests.ContainsKey(sender.IncomingIP) Then
            If Trim(MyMSGRests.Item(sender.IncomingIP)) <> "" Then
                message = Trim(MyMSGRests.Item(sender.IncomingIP)) & message
                MyMSGRests.Item(sender.IncomingIP) = ""
            End If
        End If

        '---обработка входящего сообщения
        StrArray = System.Text.RegularExpressions.Regex.Split(message, Environment.NewLine)
        For i = 0 To StrArray.Length - 1
            MyStr = StrArray(i)
            If Len(MyStr) > 0 Then
                SubStrArray = System.Text.RegularExpressions.Regex.Split(MyStr, ",")
                If i = StrArray.Length - 1 And SubStrArray.Length < 30 Then
                    '---если есть остаток - сохраняем для следующего сообщения
                    If MyMSGRests.ContainsKey(sender.IncomingIP) Then
                    Else
                        MyMSGRests.Add(sender.IncomingIP, "")
                    End If
                    MyMSGRests.Item(sender.IncomingIP) = MyStr
                End If
                If SubStrArray.Length = 30 Then
                    '-----
                    MyParam.IPAddr = sender.IncomingIP
                    MyParam.ArrEl = SubStrArray
                    '-----запуск потока для записи данных в БД
                    Dim MyDBWriteThread As New Threading.Thread(AddressOf PutLogsToDB)
                    MyDBWriteThread.IsBackground = True
                    MyDBWriteThread.Start(MyParam)
                    '-----альтернатива потоку - просто вызов ф-ции
                    'PutLogsToDB(MyParam)
                    '-----
                End If
            End If
        Next
    End Sub

    Private Sub ClientDisconnected(ByVal sender As ConnectedClient)
        '---уведомление об отсоединении клиента
        If My.Settings.MyDebug = "YES" Then
            SendMyReminder("отключение клиента TCPAvayaLogsReceiver", "отключился клиент с IP:" & vbCrLf & sender.IncomingIP)
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "ClientDisconnected 1 -> " & "Client disconnected IP = " & sender.IncomingIP)
        End If
        sender.Dispose()
        sender = Nothing
    End Sub

    Private Sub SendMyReminder(ByVal Subject As String, ByVal MyWrkString As String)
        Try
            Dim smtp As SmtpClient = New SmtpClient(My.Settings.SMTPService)
            Dim msg As New MailMessage

            msg.To.Add(My.Settings.MessageTo)
            If Trim(My.Settings.MessageCC) <> "" Then
                msg.CC.Add(My.Settings.MessageCC)
            End If
            msg.From = New MailAddress(My.Settings.MessageFrom)
            msg.Subject = Subject
            msg.Body = MyWrkString
            smtp.Send(msg)
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "SendMyReminder 1 -> " & ex.Message)
        End Try
    End Sub

    Private Sub PutLogsToDB(ByVal MyParam As Object)
        '-----надо при запуске потока для записи данных в БД
        Dim DBC As New MyDBConnector
        Dim FName As String
        Dim WrkStr As String
        Dim i As Integer

        If DBC.PutAvayaLogs(MyParam.IPAddr, MyParam.ArrEl) = False Then
            SendMyReminder("ошибка сервиса TCPAvayaLogsReceiver", "Невозможно сохранить данные в базу от " & MyParam.IPAddr & "Подробности в логах")
            '---Не положили в БД (проблемы) - сохраняем в каталог для логов
            Try
                FName = My.Settings.AvayaLogsCatalog & "\Avaya_Logs_" & Format(Now(), "yyyy-MM-dd_HH-mm-ss-fff_") & System.Guid.NewGuid.ToString & ".txt"
                Dim outfile As New StreamWriter(FName)
                WrkStr = MyParam.IPAddr & ">>"
                For i = 0 To MyParam.ArrEl.Length - 1
                    WrkStr = WrkStr & MyParam.ArrEl(i)
                    If i <> MyParam.ArrEl.Length - 1 Then
                        WrkStr = WrkStr & ","
                    End If
                Next
                outfile.Write(WrkStr & vbCrLf)
                outfile.Flush()
                outfile.Close()
                outfile = Nothing
            Catch ex As Exception
                SendMyReminder("ошибка сервиса TCPAvayaLogsReceiver", "Невозможно сохранить данные на диск. Сообщение об ошибке:" & vbCrLf & ex.Message)
                EventLog.WriteEntry("TCPAvayaLogsReceiver", "PutLogsToDB 1 -> " & ex.Message)
            End Try
        End If
        DBC.CloseMyConn()
        DBC = Nothing
        FName = Nothing
        WrkStr = Nothing
        i = Nothing
    End Sub

    Private Sub PutLogFilesToDB()
        Dim dir As New DirectoryInfo(My.Settings.AvayaLogsCatalog)
        Dim message As String
        Dim ErrFlag As Integer
        Dim StrArray() As String
        Dim StrSubArray() As String
        Dim MsgStrSubArray() As String
        Dim MyStr As String
        Dim MyIP As String
        Dim MySubstr As String
        Dim i As Integer
        Dim DBC As New MyDBConnector

        Do
            For Each MyFile As FileInfo In dir.GetFiles()
                Try
                    ErrFlag = 0
                    Using sr As New StreamReader(MyFile.FullName)
                        message = sr.ReadToEnd()
                    End Using
                    StrArray = System.Text.RegularExpressions.Regex.Split(message, Environment.NewLine)
                    For i = 0 To StrArray.Length - 1
                        MyStr = StrArray(i)
                        If Len(MyStr) > 0 Then
                            StrSubArray = System.Text.RegularExpressions.Regex.Split(MyStr, ">>")
                            MyIP = StrSubArray(0)
                            MySubstr = StrSubArray(1)
                            MsgStrSubArray = System.Text.RegularExpressions.Regex.Split(MySubstr, ",")
                            If MsgStrSubArray.Length = 30 Then
                                If DBC.PutAvayaLogs(MyIP, MsgStrSubArray) = True Then
                                Else
                                    ErrFlag = 1
                                End If
                            End If
                        End If
                    Next
                    If ErrFlag = 0 Then
                        File.Delete(MyFile.FullName)
                    End If
                Catch ex As Exception
                    EventLog.WriteEntry("TCPAvayaLogsReceiver", "PutLogsFilesToDB 1 -> " & ex.Message)
                End Try
            Next
            Thread.Sleep(My.Settings.FileReadingDelay)
        Loop
    End Sub

End Class
