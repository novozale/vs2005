Public Class MyDBConnector
    Private MyConn As ADODB.Connection                     'соединение с БД

    Public Function InitMyConn() As Boolean
        Try
            If MyConn Is Nothing Then
                MyConn = New ADODB.Connection
                MyConn.CursorLocation = 3
                MyConn.CommandTimeout = 600
                MyConn.ConnectionTimeout = 300
                MyConn.Open(My.Settings.SQLConnectionString)
            End If
            InitMyConn = True
        Catch ex As Exception
            InitMyConn = False
        End Try
    End Function

    Public Function CloseMyConn()
        Try
            MyConn.Close()
        Catch ex As Exception

        End Try
    End Function

    Public Function PutAvayaLogs(ByVal SourceIP As String, ByVal Logs() As String) As Boolean
        Dim cmd As New ADODB.Command                    'команда (spp процедура)
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        Dim MyParam9 As ADODB.Parameter                 'передаваемый параметр номер 10
        Dim MyParam10 As ADODB.Parameter                'передаваемый параметр номер 11 
        Dim MyParam11 As ADODB.Parameter                'передаваемый параметр номер 12
        Dim MyParam12 As ADODB.Parameter                'передаваемый параметр номер 13
        Dim MyParam13 As ADODB.Parameter                'передаваемый параметр номер 14 
        Dim MyParam14 As ADODB.Parameter                'передаваемый параметр номер 15
        Dim MyParam15 As ADODB.Parameter                'передаваемый параметр номер 16
        Dim MyParam16 As ADODB.Parameter                'передаваемый параметр номер 17
        Dim MyParam17 As ADODB.Parameter                'передаваемый параметр номер 18
        Dim MyParam18 As ADODB.Parameter                'передаваемый параметр номер 19
        Dim MyParam19 As ADODB.Parameter                'передаваемый параметр номер 20
        Dim MyParam20 As ADODB.Parameter                'передаваемый параметр номер 21
        Dim MyParam21 As ADODB.Parameter                'передаваемый параметр номер 22
        Dim MyParam22 As ADODB.Parameter                'передаваемый параметр номер 23
        Dim MyParam23 As ADODB.Parameter                'передаваемый параметр номер 24
        Dim MyParam24 As ADODB.Parameter                'передаваемый параметр номер 25
        Dim MyParam25 As ADODB.Parameter                'передаваемый параметр номер 26 
        Dim MyParam26 As ADODB.Parameter                'передаваемый параметр номер 27
        Dim MyParam27 As ADODB.Parameter                'передаваемый параметр номер 28
        Dim MyParam28 As ADODB.Parameter                'передаваемый параметр номер 29 
        Dim MyParam29 As ADODB.Parameter                'передаваемый параметр номер 30
        Dim MyParam30 As ADODB.Parameter                'передаваемый параметр номер 31 

        Try
            '---проверяем размерность массива
            If Logs.Length <> 30 Then
                Throw New DataException("Размер передаваемого массива не равен 30")
            End If

            If InitMyConn() = True Then
                cmd.ActiveConnection = MyConn
                cmd.CommandText = "dbo.spp_InsertData"
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmd.CommandTimeout = 300

                '----Создание параметров---------------------------------------------------
                '---0
                MyParam = cmd.CreateParameter("@AvayaIP", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 20)
                cmd.Parameters.Append(MyParam)
                MyParam.Value = SourceIP
                '---1
                MyParam1 = cmd.CreateParameter("@CALL_TIME", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam1)
                MyParam1.Value = Logs(0)
                '---2
                MyParam2 = cmd.CreateParameter("@DURATION", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
                cmd.Parameters.Append(MyParam2)
                MyParam2.Value = Logs(1)
                '---3
                MyParam3 = cmd.CreateParameter("@DURATION_S", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam3)
                If Trim(Logs(2)) = "" Then
                    MyParam3.Value = 0
                Else
                    MyParam3.Value = CInt(Logs(2))
                End If
                '---4
                MyParam4 = cmd.CreateParameter("@CALLER_PHONE", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam4)
                MyParam4.Value = Logs(3)
                '---5
                MyParam5 = cmd.CreateParameter("@DIRECTION", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 4)
                cmd.Parameters.Append(MyParam5)
                MyParam5.Value = Logs(4)
                '---6
                MyParam6 = cmd.CreateParameter("@DIALED_PHONE", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam6)
                MyParam6.Value = Logs(5)
                '---7
                MyParam7 = cmd.CreateParameter("@Field7", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam7)
                MyParam7.Value = Logs(6)
                '---8
                MyParam8 = cmd.CreateParameter("@ACC", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam8)
                MyParam8.Value = Logs(7)
                '---9
                MyParam9 = cmd.CreateParameter("@INTERNAL", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 4)
                cmd.Parameters.Append(MyParam9)
                MyParam9.Value = Logs(8)
                '---10
                MyParam10 = cmd.CreateParameter("@CallID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam10)
                MyParam10.Value = Logs(9)
                '---11
                MyParam11 = cmd.CreateParameter("@ANOTHER_RECORDS", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 4)
                cmd.Parameters.Append(MyParam11)
                MyParam11.Value = Logs(10)
                '---12
                MyParam12 = cmd.CreateParameter("@ABONENT_DEVICE1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam12)
                MyParam12.Value = Logs(11)
                '---13
                MyParam13 = cmd.CreateParameter("@ABONENT_NAME1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam13)
                MyParam13.Value = Logs(12)
                '---14
                MyParam14 = cmd.CreateParameter("@ABONENT_DEVICE2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam14)
                MyParam14.Value = Logs(13)
                '---15
                MyParam15 = cmd.CreateParameter("@ABONENT_NAME2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam15)
                MyParam15.Value = Logs(14)
                '---16
                MyParam16 = cmd.CreateParameter("@RETENTION_TIME", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam16)
                If Trim(Logs(15)) = "" Then
                    MyParam16.Value = 0
                Else
                    MyParam16.Value = CInt(Logs(15))
                End If
                '---17
                MyParam17 = cmd.CreateParameter("@PARKING_TIME", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam17)
                If Trim(Logs(16)) = "" Then
                    MyParam17.Value = 0
                Else
                    MyParam17.Value = CInt(Logs(16))
                End If
                '---18
                MyParam18 = cmd.CreateParameter("@AUTH_VALID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 4)
                cmd.Parameters.Append(MyParam18)
                MyParam18.Value = Logs(17)
                '---19
                MyParam19 = cmd.CreateParameter("@AUTH_CODE", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam19)
                MyParam19.Value = Logs(18)
                '---20
                MyParam20 = cmd.CreateParameter("@USER_PAYMENTS", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam20)
                MyParam20.Value = Logs(19)
                '---21
                MyParam21 = cmd.CreateParameter("@COST", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam21)
                MyParam21.Value = Logs(20)
                '---22
                MyParam22 = cmd.CreateParameter("@CURRENCY", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam22)
                MyParam22.Value = Logs(21)
                '---23
                MyParam23 = cmd.CreateParameter("@AOC_QTY", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam23)
                MyParam23.Value = Logs(22)
                '---24
                MyParam24 = cmd.CreateParameter("@UNITS_QTY", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam24)
                MyParam24.Value = Logs(23)
                '---25
                MyParam25 = cmd.CreateParameter("@AOC_UNITS_QTY", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam25)
                MyParam25.Value = Logs(24)
                '---26
                MyParam26 = cmd.CreateParameter("@COST_PER_UNIT", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam26)
                MyParam26.Value = Logs(25)
                '---27
                MyParam27 = cmd.CreateParameter("@ALLOWANCE", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam27)
                MyParam27.Value = Logs(26)
                '---28
                MyParam28 = cmd.CreateParameter("@OUT_CALL_REASON", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam28)
                MyParam28.Value = Logs(27)
                '---29
                MyParam29 = cmd.CreateParameter("@OUT_ABONENT_ID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam29)
                MyParam29.Value = Logs(28)
                '---30
                MyParam30 = cmd.CreateParameter("@OUT_CALL_NUMBER", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 40)
                cmd.Parameters.Append(MyParam30)
                MyParam30.Value = Logs(29)


                cmd.Execute()
                PutAvayaLogs = True
            Else
                PutAvayaLogs = False
            End If
        Catch ex As Exception
            EventLog.WriteEntry("TCPAvayaLogsReceiver", "PutAvayaLogs 1 -> " & ex.Message)
            PutAvayaLogs = False
        End Try
        cmd = Nothing
    End Function
End Class
