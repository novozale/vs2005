Public Class ConnectedClient
    Implements IDisposable
    Private disposedValue As Boolean = False        ' To detect redundant calls
    Private MyClient As System.Net.Sockets.TcpClient
    Private MyIncomingIP As String
    Private readThread As Threading.Thread
    Public Event dataReceived(ByVal sender As ConnectedClient, ByVal message As String)
    Public Event ClientDisconnected(ByVal sender As ConnectedClient)

    Sub New(ByVal client As System.Net.Sockets.TcpClient, ByVal IncomingIP As String)
        MyClient = client
        MyIncomingIP = IncomingIP
        Try
            MyClient.ReceiveTimeout = My.Settings.ReceiveTimeOut
        Catch ex As Exception
            MyClient.ReceiveTimeout = 600000 '--10 минут
        End Try

        readThread = New System.Threading.Thread(AddressOf ReadInfo)
        readThread.IsBackground = True
        readThread.Start()
    End Sub

    Private Sub ReadInfo()
        Dim readBuffer(64000) As Byte
        Dim message As String
        Dim bytesRead As Integer

        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "ReadInfo 1 -> " & "Read info Thread from " & MyIncomingIP)
        End If
        Try
            Do
                bytesRead = MyClient.GetStream.Read(readBuffer, 0, CInt(MyClient.ReceiveBufferSize))
                If bytesRead > 0 Then
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("TCPAvayaLogsReceiver", "ReadInfo 2 -> " & "Read number bytes: " & CStr(bytesRead))
                    End If
                    message = System.Text.Encoding.ASCII.GetString(readBuffer, 0, bytesRead)
                    RaiseEvent dataReceived(Me, message)
                Else
                    RaiseEvent ClientDisconnected(Me)
                    Exit Do
                End If
            Loop
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "ReadInfo 3 -> " & "Read info Error: From " & MyIncomingIP & " " & ex.Message)
            RaiseEvent ClientDisconnected(Me)
        End Try
    End Sub

    Public Property IncomingIP() As String
        Get
            Return MyIncomingIP
        End Get

        Set(ByVal value As String)
            MyIncomingIP = value
        End Set
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
                Try
                    MyClient.Close()
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("TCPAvayaLogsReceiver", "Dispose 1 -> " & "Close Client ")
                    End If
                Catch
                End Try
            End If
            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
