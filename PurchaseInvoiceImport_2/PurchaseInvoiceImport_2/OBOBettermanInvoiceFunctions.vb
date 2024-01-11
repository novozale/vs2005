Module OBOBettermanInvoiceFunctions

    Public Sub OpenOBOBettermanInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� OBO Betterman
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog
        Dim MyPurchOrder As String
        Dim MyInvoiceNumber As String
        Dim MyInvoiceDate As DateTime
        Dim MySQLStr As String                        '������� ������
        Dim i As Integer                              '������� �����

        '---��������� ����� �����
        MyDlg = New OpenFileDialog
        MyDlg.Filter = "����� Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        If MyDlg.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        '---������� �������� ���������
        Try
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        Catch ex As Exception
        End Try

        appXLSRC = CreateObject("Excel.Application")
        Try
            appXLSRC.Workbooks.Open(MyDlg.FileName)
        Catch ex As Exception
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("������ " + ex.Message)
            'Exit Sub
        End Try

        '---����������� ���� � �������� ����������
        '---������� ��������� - ����� ���� ��� ��� ���������
        '---����� ������ �� �������
        MyPurchOrder = OBOBettermanGetPurchOrderNum(appXLSRC.Worksheets(1).Range("B18").Value.ToString)
        If MyPurchOrder = "" Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("�� ��������� ���������� ����� ������ �� ������� ��������������", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        End If

        '---��� ����������
        MySQLStr = "Select SupplierCode AS Code "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & MyPurchOrder & "') AND "
        MySQLStr = MySQLStr & "(SupplierCode = N'1029') " '---���
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("��������� ������ �� ��� ������ �� �������, ��������������� ������ �� �� ������� � Scala", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Main.TextBox1.Text = declarations.MyRec.Fields("Code").Value
            trycloseMyRec()
        End If

        '---N �� ����������
        MyInvoiceNumber = OBOBettermanGetInvoiceNum(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        If MyInvoiceNumber = "" Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            MsgBox("�� ��������� ����� ����� ������� ���������� ��� ���������", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Main.textBox3.Text = MyInvoiceNumber
        End If

        '---���� �� ����������
        MyInvoiceDate = OBOBettermanGetInvoiceDate(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        If MyInvoiceDate = CDate("31/12/9999") Then
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            Main.textBox3.Text = ""
            MsgBox("�� ���������� ���� ����� ������� ���������� ��� ���������", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Main.textBox4.Text = Format(MyInvoiceDate, "dd/MM/yyyy")
        End If

        '---������ �� ����������
        If InStr(UCase(appXLSRC.Worksheets(1).Range("B13").Value.ToString), "�����") > 0 Then
            Main.textBox5.Text = 0
        Else
            Main.textBox5.Text = 12
        End If

        '---��������� - ����� ���� ��� �� ��� ���������� / �������
        MySQLStr = "SELECT COUNT(PC190300.PC19001) AS CC "
        MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC190300.PC19001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC190300.PC19012 = N'" & Main.textBox3.Text & "') AND "
        MySQLStr = MySQLStr & "(PC010300.PC01052 = N'" & MyPurchOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.Fields("CC").Value > 0 Then
            trycloseMyRec()
            Main.button3.Enabled = False
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            Main.TextBox1.Text = ""
            Main.textBox3.Text = ""
            Main.textBox4.Text = ""
            Main.textBox5.Text = ""
            Main.label6.Text = "������ �� ��� ��������� � Scala"
            Exit Sub
        Else
            trycloseMyRec()
            Main.button3.Enabled = True
            Main.label6.Text = ""
        End If

        '---���������� ��������� �������� ����
        i = 22
        While Not Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString) = "����� � ������"
            i = i + 1
        End While
        Main.progressBar1.Minimum = 0
        Main.progressBar1.Maximum = i - 23
    End Sub

    Public Sub UploadOBOBettermanInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� OBO Betterman
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String

        MyRezStr = ""

        LoadOBOBettermanInvoiceToTMPTable()
        MyRezStr = CheckUOMInOrders()
        If MyRezStr = "" Then
            If CheckEmptyInOrders() = True Then
                MyRezStr = LoadInvoiceFromTMPTable()
            Else
                MsgBox("��������� ��������� ���� � ��. �� ������ ���� ��� ������ ����������, ���� ������, ���� ���������� ����� ���� (�����������), ���� ����� ��� ��� �� ������ ����� ���� (�����������).", MsgBoxStyle.Critical, "��������!")
            End If
        End If
            UploadingRezult(MyRezStr)
    End Sub

    Private Sub LoadOBOBettermanInvoiceToTMPTable()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� OBO Betterman �� ��������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo
        Dim i As Integer                                '�������
        Dim MyInvoice As String                         '--����� ��
        Dim MyInvoiceDate As String                     '--���� ��
        Dim MyInvoiceCurrCode As Integer                '--��� ������ ��
        Dim MySalesmanCode As String                    '--��� ��������
        Dim MySalesmanName As String                    '--��� ��������
        Dim MyInvoiceCurrExchRateStr As String          '--���� ������ � ������� (������)
        Dim MyInvoiceCurrExchRate As Double             '--���� ������ � �������
        Dim MyConsPurchaseOrderNum As String            '--����� ������������������ ������ �� �������
        Dim MySupplierItemCode As String                '--��� ������ ����������
        Dim MyQTY As Double                             '--����������
        Dim MySummWithoutVAT As Double                  '--����� ��� ��� �� ������
        Dim MyCountryCode As String                     '-- ��� ������ �������������
        Dim MyCountry As String                         '-- ������ �������������
        Dim MyGTD As String                             '-- ���

        '---�������� ������ ��������� �������
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyInvoice') "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyInvoice "
        InitMyConn(False)
        declarations.MyConn.Execute(MySQLStr)

        '---�������� ����� ��������� �������
        MySQLStr = "CREATE TABLE #_MyInvoice( "
        MySQLStr = MySQLStr & "[ID] int, "                                 '--ID ������
        MySQLStr = MySQLStr & "[Invoice] [nvarchar](35), "                 '--����� ��
        MySQLStr = MySQLStr & "[InvoiceDate] [datetime], "                 '--���� ��
        MySQLStr = MySQLStr & "[InvoiceCurrCode] int, "                    '--��� ������ ��
        MySQLStr = MySQLStr & "[SalesmanCode] [nvarchar](3), "             '--��� ��������
        MySQLStr = MySQLStr & "[SalesmanName] [nvarchar](25), "            '--��� ��������
        MySQLStr = MySQLStr & "[InvoiceCurrExchRate] float, "              '--���� ������ � �������
        MySQLStr = MySQLStr & "[ConsPurchaseOrderNum] [nvarchar](10), "    '--����� ������������������ ������ �� �������
        MySQLStr = MySQLStr & "[SupplierItemCode] [nvarchar](35), "        '--��� ������ ����������
        MySQLStr = MySQLStr & "[QTY] float, "                              '--����������
        MySQLStr = MySQLStr & "[SummWithoutVAT] float, "                   '--����� ��� ��� �� ������
        MySQLStr = MySQLStr & "[Country] nvarchar(50), "                   '-- ������ �������������
        MySQLStr = MySQLStr & "[GTD] nvarchar (255), "                     '-- ���
        MySQLStr = MySQLStr & "[RestQTY] float  "                          '--������� - ���������� ����������
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        declarations.MyConn.Execute(MySQLStr)

        MyInvoice = OBOBettermanGetInvoiceNum(appXLSRC.Worksheets(1).Range("B2").Value.ToString)
        MyInvoiceDate = CStr(Format(OBOBettermanGetInvoiceDate(appXLSRC.Worksheets(1).Range("B2").Value), "dd/MM/yyyy"))
        MyInvoiceCurrCode = CInt(Main.textBox5.Text)
        MySalesmanCode = declarations.SalesmanCode
        MySalesmanName = declarations.SalesmanName
        If MyInvoiceCurrCode = 0 Then
            MyInvoiceCurrExchRate = 1
        Else
            MyInvoiceCurrExchRateStr = OBOBettermanGetInvoiceCurrExchRate(appXLSRC.Worksheets(1).Range("B14").Value.ToString)
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MyInvoiceCurrExchRate = CDbl(MyInvoiceCurrExchRateStr)
            Else
                MyInvoiceCurrExchRate = CDbl(Replace(MyInvoiceCurrExchRateStr, ",", "."))
            End If
        End If
        
        MyConsPurchaseOrderNum = OBOBettermanGetPurchOrderNum(appXLSRC.Worksheets(1).Range("B18").Value.ToString)

        i = 23
        While Not Trim(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value) = "����� � ������"
            MySupplierItemCode = OBOBettermanGetSupplierItemCode(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value)
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MyQTY = CDbl(Replace(appXLSRC.Worksheets(1).Range("K" & CStr(i)).Value.ToString, ".", ","))
            Else
                MyQTY = CDbl(appXLSRC.Worksheets(1).Range("K" & CStr(i)).Value.ToString)
            End If
            If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                MySummWithoutVAT = CDbl(Replace(appXLSRC.Worksheets(1).Range("O" & CStr(i)).Value.ToString, ".", ","))
            Else
                MySummWithoutVAT = CDbl(appXLSRC.Worksheets(1).Range("O" & CStr(i)).Value.ToString)
            End If
            If appXLSRC.Worksheets(1).Range("Y" & CStr(i)).Value = Nothing Then
                MyCountryCode = "643"
            Else
                MyCountryCode = appXLSRC.Worksheets(1).Range("Y" & CStr(i)).Value.ToString
            End If

            '---������� ������ �� ���� ������
            MySQLStr = "SELECT SY24003 "
            MySQLStr = MySQLStr & "FROM SY240300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SY24001 = N'BM') AND (SY24002 = N'" & Right("000" & MyCountryCode, 3) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                MsgBox("������ Y" & CStr(i) & " ����� ��� ������ � Scala �� ���������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Exit Sub
            Else
                MyCountry = declarations.MyRec.Fields("SY24003").Value
                trycloseMyRec()
            End If
            'MyCountry = appXLSRC.Worksheets(1).Range("X" & CStr(i)).Value.ToString

            If appXLSRC.Worksheets(1).Range("AA" & CStr(i)).Value = Nothing Then
                MyGTD = ""
            Else
                MyGTD = Trim(appXLSRC.Worksheets(1).Range("AA" & CStr(i)).Value.ToString)
            End If

            MySQLStr = "INSERT INTO #_MyInvoice "
            MySQLStr = MySQLStr & "(ID, Invoice, InvoiceDate, InvoiceCurrCode, SalesmanCode, SalesmanName, InvoiceCurrExchRate, "
            MySQLStr = MySQLStr & "ConsPurchaseOrderNum, SupplierItemCode, QTY, SummWithoutVAT, Country, GTD, RestQTY) "
            MySQLStr = MySQLStr & "VALUES (" & CStr(i - 21) & ", "
            MySQLStr = MySQLStr & "N'" & MyInvoice & "', "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & MyInvoiceDate & "', 103), "
            MySQLStr = MySQLStr & CStr(MyInvoiceCurrCode) & ", "
            MySQLStr = MySQLStr & "N'" & MySalesmanCode & "', "
            MySQLStr = MySQLStr & "N'" & MySalesmanName & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyInvoiceCurrExchRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "N'" & MyConsPurchaseOrderNum & "', "
            MySQLStr = MySQLStr & "N'" & MySupplierItemCode & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(MySummWithoutVAT), ",", ".") & ", "
            MySQLStr = MySQLStr & "N'" & MyCountry & "', "
            MySQLStr = MySQLStr & "N'" & MyGTD & "', "
            MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
            InitMyConn(False)
            declarations.MyConn.Execute(MySQLStr)

            i = i + 1
        End While

        'MySQLStr = "SELECT * FROM #_MyInvoice "
        'InitMyConn(False)
        'InitMyRec(False, MySQLStr)
        'While declarations.MyRec.EOF <> True
        ' MsgBox("ID:" & declarations.MyRec.Fields("ID").Value.ToString & Chr(13) & " Invoice:" & declarations.MyRec.Fields("Invoice").Value.ToString & Chr(13) & " InvoiceDate:" & declarations.MyRec.Fields("InvoiceDate").Value.ToString & Chr(13) & " InvoiceCurrCode:" & declarations.MyRec.Fields("InvoiceCurrCode").Value.ToString & Chr(13) & " SalesmanCode:" & declarations.MyRec.Fields("SalesmanCode").Value.ToString & Chr(13) & " SalesmanName:" & declarations.MyRec.Fields("SalesmanName").Value.ToString & Chr(13) & " InvoiceCurrExchRate:" & declarations.MyRec.Fields("InvoiceCurrExchRate").Value.ToString & Chr(13) & " ConsPurchaseOrderNum:" & declarations.MyRec.Fields("ConsPurchaseOrderNum").Value.ToString & Chr(13) & " SupplierItemCode:" & declarations.MyRec.Fields("SupplierItemCode").Value.ToString & Chr(13) & " QTY:" & declarations.MyRec.Fields("QTY").Value.ToString & Chr(13) & " SummWithoutVAT:" & declarations.MyRec.Fields("SummWithoutVAT").Value.ToString & Chr(13) & " Country:" & declarations.MyRec.Fields("Country").Value.ToString & Chr(13) & " GTD:" & declarations.MyRec.Fields("GTD").Value.ToString & Chr(13) & " RestQTY:" & declarations.MyRec.Fields("RestQTY").Value.ToString, MsgBoxStyle.Information, "��������!")
        'declarations.MyRec.MoveNext()
        'End While

    End Sub

    Private Function OBOBettermanGetPurchOrderNum(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ ������ ������ �� ������ �� ��� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        '������� ���������

        'MyStr = Mid(MyStr, 9)
        'If Len(MyStr) < 9 Then
        '    OBOBettermanGetPurchOrderNum = ""
        'Else
        '    MyPos = InStr(MyStr, " ")
        '    If MyPos = 0 Then
        '        OBOBettermanGetPurchOrderNum = ""
        '    Else
        '        MyStr = Mid(MyStr, 1, MyPos - 1)
        '        OBOBettermanGetPurchOrderNum = Right("0000000000" & MyStr, 10)
        '    End If
        'End If
        '---����� ��������� ����� �� ������� - �� � N, �� ������, �� ����� - ������������� � ����, ��� ��� ��� ���������� � 07
        MyPos = InStr(MyStr, "07")
        If MyPos = 0 Then
            OBOBettermanGetPurchOrderNum = ""
        Else
            MyStr = Mid(MyStr, MyPos, 10)
            OBOBettermanGetPurchOrderNum = Right("0000000000" & MyStr, 10)
        End If
    End Function

    Private Function OBOBettermanGetInvoiceNum(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ �� �� ������ �� ��� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        '������� ���������

        MyStr = Mid(MyStr, 16)
        If Len(MyStr) < 1 Then
            OBOBettermanGetInvoiceNum = ""
        Else
            MyPos = InStr(MyStr, " ")
            If MyPos = 0 Then
                OBOBettermanGetInvoiceNum = ""
            Else
                MyStr = Mid(MyStr, 1, MyPos - 1)
                OBOBettermanGetInvoiceNum = MyStr
            End If
        End If
    End Function

    Private Function OBOBettermanGetInvoiceDate(ByVal MyStr As String) As DateTime
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� �� �� ������ �� ��� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        '������� ���������

        MyStr = Mid(MyStr, 16)
        If Len(MyStr) < 1 Then
            OBOBettermanGetInvoiceDate = CDate("31/12/9999")
        Else
            MyPos = InStr(MyStr, " ")
            If MyPos = 0 Then
                OBOBettermanGetInvoiceDate = CDate("31/12/9999")
            Else
                MyStr = Mid(MyStr, MyPos + 1)
                If Len(MyStr) < 1 Then
                    OBOBettermanGetInvoiceDate = CDate("31/12/9999")
                Else
                    MyPos = InStr(MyStr, " ")
                    If MyPos = 0 Then
                        OBOBettermanGetInvoiceDate = CDate("31/12/9999")
                    Else
                        MyStr = Mid(MyStr, MyPos + 1)
                        OBOBettermanGetInvoiceDate = CDate(MyStr)
                    End If
                End If
            End If
        End If
    End Function

    Private Function OBOBettermanGetInvoiceCurrExchRate(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������ ������ �� ������ �� ��� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        '������� ���������

        MyPos = InStr(MyStr, "����")
        If MyPos = 0 Then
            OBOBettermanGetInvoiceCurrExchRate = "0"
        Else
            MyStr = Mid(MyStr, MyPos + 5)
            OBOBettermanGetInvoiceCurrExchRate = MyStr
        End If
    End Function

    Private Function OBOBettermanGetSupplierItemCode(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� ������ ���������� �� ������ �� ��� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyPos As Double                        '������� ���������
        Dim subStr As String

        MyPos = InStr(MyStr, "���:")
        If MyPos = 0 Then
            OBOBettermanGetSupplierItemCode = Trim(MyStr)
        Else
            subStr = MyStr.Substring(MyPos + 4)
            OBOBettermanGetSupplierItemCode = Trim(subStr)
        End If
    End Function
End Module
