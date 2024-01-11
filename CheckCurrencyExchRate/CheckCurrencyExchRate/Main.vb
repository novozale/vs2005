Imports System.Xml

Module Main

    Sub Main()
        '///////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сверка курсов ЦБ РФ с курсами, выставленными в Scala
        '//
        '///////////////////////////////////////////////////////////////////////////////////////

        GetScalaCurrExchRate()

        GetCBRFCurrExchRate()

        If (EURCBRF = 0 Or USDCBRF = 0 Or GBPCBRF = 0 Or CNYCBRF = 0 Or TRYCBRF = 0 Or KZTCBRF = 0) Or _
            (USDScala1 = 0 Or USDScala2 = 0 Or EURScala1 = 0 Or EURScala2 = 0 Or GBPScala1 = 0 Or GBPScala2 = 0 Or _
            CNYScala1 = 0 Or CNYScala2 = 0 Or TRYScala1 = 0 Or TRYScala2 = 0 Or KZTScala1 = 0 Or KZTScala2 = 0) Or _
            (EURCBRF <> EURScala1 Or EURCBRF <> EURScala2) Or _
            (USDCBRF <> USDScala1 Or USDCBRF <> USDScala2) Or _
            (GBPCBRF <> GBPScala1 Or GBPCBRF <> GBPScala2) Or _
            (CNYCBRF <> CNYScala1 Or CNYCBRF <> CNYScala2) Or _
            (TRYCBRF <> TRYScala1 Or TRYCBRF <> TRYScala2) Or _
            (KZTCBRF <> KZTScala1 Or KZTCBRF <> KZTScala2) Then
            SendErrMessage()
        End If

    End Sub

    Private Sub GetScalaCurrExchRate()
        '///////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение курсов валюты из Scala на текущую дату
        '//
        '///////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-------Доллары
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 1) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            USDScala1 = 0
            USDScala2 = 0
        Else
            USDScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 5)
            USDScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 5)
            trycloseMyRec()
        End If

        '-------Евро
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 12) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            EURScala1 = 0
            EURScala2 = 0
        Else
            EURScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 5)
            EURScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 5)
            trycloseMyRec()
        End If

        '-------Фунты стерлингов
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 4) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            GBPScala1 = 0
            GBPScala2 = 0
        Else
            GBPScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 5)
            GBPScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 5)
            trycloseMyRec()
        End If

        '-------юань
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 6) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            CNYScala1 = 0
            CNYScala2 = 0
        Else
            CNYScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 5)
            CNYScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 5)
            trycloseMyRec()
        End If

        '-------турецкая лира
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 13) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            TRYScala1 = 0
            TRYScala2 = 0
        Else
            TRYScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 5)
            TRYScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 5)
            trycloseMyRec()
        End If

        '-------казахстанский тенге
        MySQLStr = "SELECT SYCH006, SYCH007 "
        MySQLStr = MySQLStr & "FROM SYCH0100 "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = 11) AND (SYCH004 <= Dateadd(dd,1,GETDATE())) AND (SYCH005 > Dateadd(dd,1,GETDATE())) " '----проверяем завтрашний день
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            KZTScala1 = 0
            KZTScala2 = 0
        Else
            KZTScala1 = Math.Round(Declarations.MyRec.Fields("SYCH006").Value, 6)
            KZTScala2 = Math.Round(Declarations.MyRec.Fields("SYCH007").Value, 6)
            trycloseMyRec()
        End If

    End Sub

    Private Sub GetCBRFCurrExchRate()
        '///////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение курсов валюты из ЦБ РФ на текущую дату
        '//
        '///////////////////////////////////////////////////////////////////////////////////////
        Dim sURI As String
        Dim oHttp As Object
        Dim htmlcode, outstrD, outstrE, outstrG As String
        Dim inpdate As Date
        Dim d, m, y As Integer
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer

        inpdate = DateAdd(DateInterval.Day, 1, Now()) '----проверяем завтрашний день
        d = DatePart("d", inpdate)
        m = DatePart("m", inpdate)
        y = DatePart("yyyy", inpdate)
        'sURI = "http://cbr.ru/currency_base/daily.aspx?C_month=" & Right("00" & CStr(m), 2) & "&C_year=" _
        '          & y & "&date_req=" & Right("00" & CStr(d), 2) & "%2F" & Right("00" & CStr(m), 2) & "%2F" & y
        'sURI = "https://cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To=" & Right("00" & CStr(d), 2) & "." & Right("00" & CStr(m), 2) & "." & y
        sURI = "http://cbr.ru/scripts/XML_daily.asp?date_req=" & Right("00" & CStr(d), 2) & "." & Right("00" & CStr(m), 2) & "." & y
        Try
            'oHttp = CreateObject("MSXML2.XMLHTTP")
            'If Err.Number <> 0 Then
            'oHttp = CreateObject("MSXML.XMLHTTPRequest")
            'End If

            oHttp = CreateObject("MSXML2.ServerXMLHTTP")

            If oHttp Is Nothing Then
                EURCBRF = 0
                USDCBRF = 0
                GBPCBRF = 0
                CNYCBRF = 0
                TRYCBRF = 0
                KZTCBRF = 0
                Exit Sub
            End If

            oHttp.setTimeouts(120000, 120000, 120000, 120000)

            oHttp.Open("GET", sURI, False)
            oHttp.Send()
            htmlcode = oHttp.responseText
        Catch
            USDCBRF = 0
            EURCBRF = 0
            GBPCBRF = 0
            CNYCBRF = 0
            TRYCBRF = 0
            KZTCBRF = 0
        End Try

        Try
            xmldoc.LoadXml(htmlcode)
            xmlnode = xmldoc.GetElementsByTagName("Valute")
            For i = 0 To xmlnode.Count - 1
                'MsgBox(xmlnode(i).ChildNodes(1).InnerText + " --- " + xmlnode(i).ChildNodes(4).InnerText, MsgBoxStyle.Information, "Внимание!")
                '----------Доллар
                If xmlnode(i).ChildNodes(1).InnerText.Equals("USD") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        USDCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    Else
                        USDCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    End If
                End If
                '----------Евро
                If xmlnode(i).ChildNodes(1).InnerText.Equals("EUR") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        EURCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    Else
                        EURCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    End If
                End If
                '----------Фунты стерлингов
                If xmlnode(i).ChildNodes(1).InnerText.Equals("GBP") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        GBPCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    Else
                        GBPCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    End If
                End If
                '----------юани
                If xmlnode(i).ChildNodes(1).InnerText.Equals("CNY") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        CNYCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    Else
                        CNYCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    End If
                End If
                '----------турецкие лиры
                If xmlnode(i).ChildNodes(1).InnerText.Equals("TRY") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        TRYCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    Else
                        TRYCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 5)
                    End If
                End If
                '----------казахстанские тенге
                If xmlnode(i).ChildNodes(1).InnerText.Equals("KZT") Then
                    If aa.CurrentInfo.NumberDecimalSeparator = "." Then
                        KZTCBRF = Math.Round(CDbl(Replace(xmlnode(i).ChildNodes(4).InnerText, ",", ".")) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 6)
                    Else
                        KZTCBRF = Math.Round(CDbl(xmlnode(i).ChildNodes(4).InnerText) / CDbl(xmlnode(i).ChildNodes(2).InnerText), 6)
                    End If
                End If
            Next

            ''----------Доллар
            'outstrD = Mid(htmlcode, InStr(1, htmlcode, "USD") + 47 + 18, 7)
            'If Char.IsDigit(Right(outstrD, 1)) = False Then
            '    outstrD = Left(outstrD, 7)
            'End If
            'If aa.CurrentInfo.NumberDecimalSeparator = "." Then
            '    USDCBRF = CDbl(Replace(outstrD, ",", "."))
            'Else
            '    USDCBRF = CDbl(outstrD)
            'End If
            ''----------Евро
            'outstrE = Mid(htmlcode, InStr(1, htmlcode, "EUR") + 41 + 18, 7)
            'If Char.IsDigit(Right(outstrE, 1)) = False Then
            '    outstrE = Left(outstrE, 7)
            'End If
            'If aa.CurrentInfo.NumberDecimalSeparator = "." Then
            '    EURCBRF = CDbl(Replace(outstrE, ",", "."))
            'Else
            '    EURCBRF = CDbl(outstrE)
            'End If
            ''----------Фунты стерлингов
            'outstrG = Mid(htmlcode, InStr(1, htmlcode, "GBP") + 77 + 18, 7)
            'If Char.IsDigit(Right(outstrG, 1)) = False Then
            '    outstrG = Left(outstrG, 7)
            'End If
            'If aa.CurrentInfo.NumberDecimalSeparator = "." Then
            '    GBPCBRF = CDbl(Replace(outstrG, ",", "."))
            'Else
            '    GBPCBRF = CDbl(outstrG)
            'End If
        Catch
            USDCBRF = 0
            EURCBRF = 0
            GBPCBRF = 0
            CNYCBRF = 0
            TRYCBRF = 0
            KZTCBRF = 0
        End Try
    End Sub

    Private Sub SendErrMessage()
        '///////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка сообщения о проблемах с выставлением курсов валют
        '//
        '///////////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage

        msg.To.Add(My.Settings.AddressITD)
        msg.To.Add(My.Settings.AddressITM)
        msg.To.Add(My.Settings.AddressCFO)
        msg.To.Add(My.Settings.AddressCA)
        msg.To.Add(My.Settings.AddressACC1)
        msg.To.Add(My.Settings.AddressACC2)
        msg.To.Add(My.Settings.AddressGUT1)
        msg.To.Add(My.Settings.AddressGUT2)
        msg.To.Add(My.Settings.AddressWHM1)
        msg.To.Add(My.Settings.AddressWHM2)
        msg.To.Add(My.Settings.AddressWHM3)
        msg.To.Add(My.Settings.AddressWHM4)
        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "Проблема с выставлением курсов валют в Scala"
        MyMsgStr = "Проблема с выставлением курсов валют в Scala" & Chr(13) & Chr(13)

        MyMsgStr = MyMsgStr & "курс доллара ЦБ РФ: " & CStr(USDCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы доллара в Scala: " & CStr(USDScala1) & " и " & CStr(USDScala2) & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "курс евро ЦБ РФ: " & CStr(EURCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы евро в Scala: " & CStr(EURScala1) & " и " & CStr(EURScala2) & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "курс фунта стерлингов ЦБ РФ: " & CStr(GBPCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы фунта стерлингов в Scala: " & CStr(GBPScala1) & " и " & CStr(GBPScala2) & Chr(13) & Chr(13)

        MyMsgStr = MyMsgStr & "курс юаня ЦБ РФ: " & CStr(CNYCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы юаня в Scala: " & CStr(CNYScala1) & " и " & CStr(CNYScala2) & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "курс турецкой лиры ЦБ РФ: " & CStr(TRYCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы турецкой лиры в Scala: " & CStr(TRYScala1) & " и " & CStr(TRYScala2) & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "курс казахстанского тенге ЦБ РФ: " & CStr(KZTCBRF) & Chr(13)
        MyMsgStr = MyMsgStr & "Курсы казахстанского тенге в Scala: " & CStr(KZTScala1) & " и " & CStr(KZTScala2) & Chr(13) & Chr(13)

        MyMsgStr = MyMsgStr & "Проверьте выставленные в Scala курсы и доступность сайта центробанка " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "________________________" & Chr(13)
        MyMsgStr = MyMsgStr & "На письмо не отвечайте, отправлено автоматически." & Chr(13)
        msg.Body = MyMsgStr
        smtp.Send(msg)
    End Sub
End Module
