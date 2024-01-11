Public Class SendInfo

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SendInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ���������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        TextBox1.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
        If MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "���������" Then
            ComboBox1.SelectedText = "���������"
        ElseIf MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "�������� � ������� ��������" Then
            ComboBox1.SelectedText = "�������� � ������� ��������"
        Else
            ComboBox1.SelectedText = "��������"
        End If
        TextBox3.Text = Declarations.MyCustomerCode
        '-----EMail (� ����������� �� ������������)
        If MyEmail = 1 Then
            MySQLStr = "SELECT tbl_CRM_Contacts.ContactEMail AS EMail "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Contacts.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Companies.ScalaCustomerCode = N'" & Trim(Declarations.MyCustomerCode) & "') "
            MySQLStr = MySQLStr & "AND ('���������� ����: ' + tbl_CRM_Contacts.ContactName + ' �������: ' + tbl_CRM_Contacts.ContactPhone + ' E-Mail: ' + tbl_CRM_Contacts.ContactEMail "
            MySQLStr = MySQLStr & "= N'" & Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                TextBox2.Text = ""
                trycloseMyRec()
            Else
                TextBox2.Text = Declarations.MyRec.Fields("EMail").Value
                trycloseMyRec()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDel As Integer
        Dim MySQLStr As String

        If CheckFields() = True Then
            If ComboBox1.SelectedText = "���������" Then
                MyDel = 0
            ElseIf ComboBox1.SelectedText = "�������� � ������� ��������" Then
                MyDel = 2
            Else
                MyDel = 1
            End If
            SendInfoByEmail(TextBox1.Text, MyDel, TextBox2.Text, MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(12).Value)
            MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Info "
            MySQLStr = MySQLStr & "SET IsReminderSend = 1 "
            MySQLStr = MySQLStr & "WHERE (ID = " & MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString & ")"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(14).Value = 1
            Me.Close()
            MsgBox("����������� ������� ����������.", MsgBoxStyle.Information, "��������!")
        End If
    End Sub

    Public Function CheckFields() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim email As New System.Text.RegularExpressions.Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")

        If Trim(TextBox2.Text) = "" Then
            MsgBox("� ���� ""�������� ������ ��������"" ������ ���� ������� ���� �� ���� �����, �� �������� ����� ���������� �����������.", MsgBoxStyle.Critical, "��������!")
            TextBox2.Select()
            Return False
        End If

        If email.IsMatch(TextBox2.Text) Then
        Else
            MsgBox("� ���� ""�������� ������ ��������"" ������ ���� ������� ���������� �����, �� �������� ����� ���������� �����������.", MsgBoxStyle.Critical, "��������!")
            TextBox2.Select()
            Return False
        End If

        Return True
    End Function

    Private Sub SendInfoByEmail(ByVal DeliveryID As Integer, ByVal DeliveryOrNot As Integer, ByVal MyEMail As String, ByVal RequestedDate As DateTime)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String
        Dim MySQLStr As String
        Dim AttachmentsFlag As Integer
        Dim MyCatalog As String
        Dim MyOrder As String
        Dim MyEMailArr() As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        MyCatalog = "c:\Data_Exchange\" + Declarations.MyCustomerCode

        MyEMailArr = MyEMail.Split(";".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For y As Integer = 0 To MyEMailArr.Length - 1
            msg.To.Add(MyEMailArr(y))
        Next

        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "����������� �� �������� �������� � ��������� �������"
        MyMsgStr = "��������� ������!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "    �� ���������� ��� �� ����� ����� �������� � �������� ���������� ����������������� � �������������������." & Chr(13)
        If DeliveryOrNot = 0 Then
            MyMsgStr = MyMsgStr + "���������� ��� � ���������� ����� ������� � ���������� � ������ ������. " & Chr(13)
            MyMsgStr = MyMsgStr + "���� ���������� �������: " & Microsoft.VisualBasic.Right("00" & CStr(RequestedDate.Day), 2) & "\" & Microsoft.VisualBasic.Right("00" & CStr(RequestedDate.Month), 2) & "\" & RequestedDate.Year & Chr(13)
            MyMsgStr = MyMsgStr + "���������� � ������� ������������ ��������� � �������������� � ������ ������ (�����)" & Chr(13) & Chr(13)
        Else
            MyMsgStr = MyMsgStr + "���������� ��� � ����������� �������� ����� ������� � ������ ������. " & Chr(13)
            MyMsgStr = MyMsgStr + "���� ��������������� ��������: " & Microsoft.VisualBasic.Right("00" & CStr(RequestedDate.Day), 2) & "\" & Microsoft.VisualBasic.Right("00" & CStr(RequestedDate.Month), 2) & "\" & RequestedDate.Year & Chr(13)
            MyMsgStr = MyMsgStr + "���������� �� ����������� ������������ ��������� � �������������� � ������ ������ (�����)" & Chr(13) & Chr(13)
        End If
        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "� ���������," & Chr(13)
        MyMsgStr = MyMsgStr + "��� ""��������"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. �� ������ ������� �� ��������, ��� �������������� ��������. "
        msg.Body = MyMsgStr

        '----------����������--------
        AttachmentsFlag = 0
        MySQLStr = "SELECT OrderNum "
        MySQLStr = MySQLStr & "FROM tbl_Shipments_SalesmanWP_Details "
        MySQLStr = MySQLStr & "WHERE (ShipmentsID = " & DeliveryID & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                MyOrder = Declarations.MyRec.Fields("OrderNum").Value
                GetOrderReport(MyCatalog, MyOrder)
                If FileExists(MyCatalog + "\" + MyOrder + ".pdf") Then
                    Dim MyData As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(MyCatalog + "\" + MyOrder + "." + My.Settings.Format, System.Net.Mime.MediaTypeNames.Application.Octet)
                    msg.Attachments.Add(MyData)
                    AttachmentsFlag = AttachmentsFlag + 1
                End If
                Declarations.MyRec.MoveNext()
            End While
        End If
        trycloseMyRec()

        If AttachmentsFlag <> 0 Then
            smtp.Send(msg)
            For Each a As System.Net.Mail.Attachment In msg.Attachments
                a.Dispose()
            Next
            msg = Nothing
            smtp = Nothing
        End If
        Try
            Dim di As New IO.DirectoryInfo(MyCatalog)
            di.Delete(True)
        Catch ex As Exception
        End Try
    End Sub

    Private Function FileExists(ByVal FileFullPath As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ���������� ���� ��� ��� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If Trim(FileFullPath) = "" Then Return False

        Dim f As New IO.FileInfo(FileFullPath)
        Return f.Exists

    End Function

    Private Sub GetOrderReport(ByVal MyCatalog As String, ByVal MyOrder As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ � ���� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyFile As IO.FileStream
        Dim execInfo As New spbprd5.ExecutionInfo

        If (Not System.IO.Directory.Exists(MyCatalog)) Then
            System.IO.Directory.CreateDirectory(MyCatalog)
        End If

        Dim rs As New spbprd5.ReportExecutionService
        rs.Credentials = System.Net.CredentialCache.DefaultCredentials
        rs.Url = My.Settings.ShipmentsWorkplace_spbprd5_ReportExecutionService

        ' Render arguments.
        Dim result As Byte() = Nothing
        Dim reportPath As String = My.Settings.MyReport
        Dim format As String = My.Settings.Format
        Dim historyID As String = Nothing

        ' Prepare report parameter.
        Dim parameters(0) As spbprd5.ParameterValue
        parameters(0) = New spbprd5.ParameterValue()
        parameters(0).Name = "MyOrderNumber"
        parameters(0).Value = Trim(MyOrder)

        Dim encoding As String = String.Empty
        Dim mimeType As String = String.Empty
        Dim warnings As spbprd5.Warning() = Nothing
        Dim streamIDs As String() = Nothing
        Dim deviceInfo As String = Nothing
        Dim Extencion As String = Nothing
        Dim MyLng As String = "ru-RU"
        execInfo = rs.LoadReport(reportPath, historyID)
        rs.SetExecutionParameters(parameters, MyLng)
        rs.Timeout = -1
        result = rs.Render(format, deviceInfo, Extencion, mimeType, encoding, warnings, streamIDs)

        Using stream As IO.FileStream = IO.File.OpenWrite(MyCatalog + "\" + MyOrder + "." + My.Settings.Format)
            stream.Write(result, 0, result.Length)
        End Using
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��������� �� CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MyContactInfo = New ContactInfo
        MyContactInfo.StartParam = "EMail"
        MyContactInfo.ShowDialog()
    End Sub
End Class