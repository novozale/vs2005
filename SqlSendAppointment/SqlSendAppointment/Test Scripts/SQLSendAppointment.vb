Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server

Imports System.Net.Mail
Imports System.Text

Partial Public Class UserDefinedFunctions
    <Microsoft.SqlServer.Server.SqlFunction()> _
    Public Shared Function SQLSendAppointment(ByVal from As String, ByVal recepient As String, ByVal subject As String, ByVal MyDate As String, _
        ByVal MyGUID As String, ByVal MyCompany As String, ByVal MyContact As String, ByVal MySMTP As String, ByVal MyPort As Integer) As SqlString
        ' ///////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка напоминания для календаря
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim mMailMessage As New MailMessage()

        Try
            mMailMessage.From = New MailAddress(from)
            mMailMessage.To.Add(New MailAddress(recepient))

            mMailMessage.Subject = subject
            mMailMessage.Priority = MailPriority.Normal

            Dim str As StringBuilder
            str = New StringBuilder
            'str.AppendLine("BEGIN:VCALENDAR")
            'str.AppendLine("PRODID:-//Skandika")
            'str.AppendLine("VERSION:2.0")
            'str.AppendLine("METHOD:REQUEST")
            'str.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE")
            'str.AppendLine("BEGIN:VEVENT")
            'str.AppendLine("ATTENDEE;RSVP=TRUE;ROLE=REQ-PARTICIPANT;CUTYPE=GROUP:MAILTO:" & recepient)
            'str.AppendLine("STATUS:CONFIRMED")
            'str.AppendLine("DTSTART:" & MyDate & "T050000Z")
            'str.AppendLine("DTSTAMP:" & MyDate & "T050000Z")
            'str.AppendLine("DTEND:" & MyDate & "T133000Z")
            'str.AppendLine("SEQUENCE:0")
            'str.AppendLine("LOCATION: Russia")
            ''str.AppendLine(String.Format("UID:{0}", Guid.NewGuid()))
            'str.AppendLine("UID:" & MyGUID)
            'str.AppendLine("CLASS:PUBLIC")
            'str.AppendLine(String.Format("DESCRIPTION:{0}", MyCompany & " " & MyContact))
            'str.AppendLine(String.Format("SUMMARY:{0}", mMailMessage.Subject))
            'str.AppendLine("BEGIN:VALARM")
            'str.AppendLine("TRIGGER:-PT15M")
            'str.AppendLine("ACTION:DISPLAY")
            'str.AppendLine("DESCRIPTION:Reminder")
            'str.AppendLine("END:VALARM")
            'str.AppendLine("END:VEVENT")
            'str.AppendLine("END:VCALENDAR")

            'str.AppendLine("BEGIN:VCALENDAR")
            'str.AppendLine("PRODID:-//Mozilla.org/NONSGML Mozilla Calendar V1.1//EN")
            'str.AppendLine("VERSION:2.0")
            'str.AppendLine("BEGIN:VTIMEZONE")
            'str.AppendLine("TZID:Europe/Moscow")
            'str.AppendLine("BEGIN:STANDARD")
            'str.AppendLine("TZOFFSETFROM:+0300")
            'str.AppendLine("TZOFFSETTO:+0300")
            'str.AppendLine("TZNAME:MSK")
            'str.AppendLine("DTSTART:19700101T000000")
            'str.AppendLine("END:STANDARD")
            'str.AppendLine("END:VTIMEZONE")

            'str.AppendLine("BEGIN:VEVENT")
            'str.AppendLine("LAST-MODIFIED:" & MyDate & "T060000Z")
            'str.AppendLine("DTSTAMP:" & MyDate & "T060000Z")
            'str.AppendLine("UID:" & MyGUID)
            ''str.AppendLine(String.Format("SUMMARY:{0}", mMailMessage.Subject))
            'str.AppendLine("SUMMARY:тестовое событие")
            'str.AppendLine("STATUS:CONFIRMED")
            ''str.AppendLine("ORGANIZER;SENT-BY=""" & recepient & """:mailto:" & from)
            ''str.AppendLine("ORGANIZER;SENT-BY=" & recepient & ":mailto:" & from)
            ''str.AppendLine("ORGANIZER:mailto:" & from)
            ''str.AppendLine("ATTENDEE;PARTSTAT=ACCEPTED;CUTYPE=INDIVIDUAL:mailto:" & recepient)
            'str.AppendLine("DTSTART;TZID=Europe/Moscow:" & MyDate & "T060000Z")
            'str.AppendLine("DTEND;TZID=Europe/Moscow:" & MyDate & "T143000Z")
            ''str.AppendLine(String.Format("DESCRIPTION:{0}", MyCompany & " " & MyContact))
            'str.AppendLine("DESCRIPTION:тестовое описание")
            'str.AppendLine("LOCATION:Тестовое место")
            ''str.AppendLine("X-MOZ-GENERATION:1")
            ''str.AppendLine("X-MS-OLK-SENDER:mailto:" & recepient)
            ''str.AppendLine("X-MOZ-INVITED-ATTENDEE:mailto:" & recepient)
            'str.AppendLine("CLASS:PUBLIC")
            'str.AppendLine("TRANSP:OPAQUE")
            'str.AppendLine("SEQUENCE:0")
            'str.AppendLine("BEGIN:VALARM")
            'str.AppendLine("ACTION:DISPLAY")
            'str.AppendLine("TRIGGER;VALUE=DURATION:-PT15M")
            'str.AppendLine("DESCRIPTION:Reminder")
            'str.AppendLine("END:VALARM")
            ''str.AppendLine("BEGIN:VALARM")
            ''str.AppendLine("TRIGGER:-PT15M")
            ''str.AppendLine("ACTION:DISPLAY")
            ''str.AppendLine("DESCRIPTION:Reminder")
            ''str.AppendLine("END:VALARM")
            'str.AppendLine("END:VEVENT")

            'str.AppendLine("END:VCALENDAR")

            str.AppendLine("BEGIN:VCALENDAR")
            str.AppendLine("PRODID:-//Skandika")
            str.AppendLine("VERSION:2.0")
            str.AppendLine("BEGIN:VTIMEZONE")
            str.AppendLine("TZID:Europe/Moscow")
            str.AppendLine("BEGIN:STANDARD")
            str.AppendLine("TZOFFSETFROM:+0300")
            str.AppendLine("TZOFFSETTO:+0300")
            str.AppendLine("TZNAME:MSK")
            str.AppendLine("DTSTART:19700101T000000")
            str.AppendLine("END:STANDARD")
            str.AppendLine("END:VTIMEZONE")
            str.AppendLine("METHOD:REQUEST")
            str.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE")

            'str.AppendLine("BEGIN:VEVENT")
            'str.AppendLine("STATUS:CONFIRMED")
            'str.AppendLine("ORGANIZER:MAILTO:" & recepient)
            ''str.AppendLine("ATTENDEE;RSVP=TRUE;ROLE=REQ-PARTICIPANT;CUTYPE=GROUP:MAILTO:" & recepient)
            'str.AppendLine("ATTENDEE;MAILTO:" & recepient)
            'str.AppendLine("DTSTART;TZID=Europe/Moscow:" & MyDate & "T090000")
            'str.AppendLine("DTEND;TZID=Europe/Moscow:" & MyDate & "T173000")
            'str.AppendLine(String.Format("SUMMARY:{0}", mMailMessage.Subject))
            'str.AppendLine(String.Format("DESCRIPTION:{0}", MyCompany & " " & MyContact))
            'str.AppendLine("LOCATION:Russia")
            'str.AppendLine("BEGIN:VALARM")
            'str.AppendLine("TRIGGER:-PT15M")
            'str.AppendLine("ACTION:DISPLAY")
            'str.AppendLine("DESCRIPTION:Reminder")
            'str.AppendLine("END:VALARM")
            'str.AppendLine("END:VEVENT")
            str.AppendLine("BEGIN:VEVENT")
            str.AppendLine("ORGANIZER:MAILTO:" & recepient)
            'str.AppendLine("ATTENDEE:MAILTO:alexander.novozhilov@skandikagroup.ru")
            str.AppendLine("STATUS:CONFIRMED")
            str.AppendLine("DTSTART;TZID=Europe/Moscow:" & MyDate & "T090000")
            'str.AppendLine("DTSTAMP:20230712T050000Z")
            str.AppendLine("DTEND;TZID=Europe/Moscow:" & MyDate & "T173000")
            'str.AppendLine("SEQUENCE:0")
            str.AppendLine("LOCATION: Russia")
            'str.AppendLine("CLASS:PUBLIC")
            str.AppendLine(String.Format("DESCRIPTION:{0}", MyCompany & " " & MyContact))
            str.AppendLine(String.Format("SUMMARY:{0}", mMailMessage.Subject))
            str.AppendLine("BEGIN:VALARM")
            str.AppendLine("TRIGGER:-PT15M")
            str.AppendLine("ACTION:DISPLAY")
            str.AppendLine("DESCRIPTION:Reminder")
            str.AppendLine("END:VALARM")
            str.AppendLine("END:VEVENT")

            str.AppendLine("END:VCALENDAR")

            Dim ct As System.Net.Mime.ContentType
            ct = New System.Net.Mime.ContentType("text/calendar")
            ct.Parameters.Add("method", "REQUEST")
            ct.Parameters.Add("name", "event.ics")

            Dim avCal As AlternateView
            avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct)
            mMailMessage.AlternateViews.Add(avCal)

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient(MySMTP, MyPort)

            mSmtpClient.Credentials = New System.Net.NetworkCredential()

            ' Send the mail message
            mSmtpClient.Send(mMailMessage)
        Catch

        End Try
        Return New SqlString("Send")
    End Function
End Class
