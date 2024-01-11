Module InvoiceFunctions

    Public Sub OpenInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Main.ComboBox1.SelectedItem = "���" Then
            OpenABBInvoiceFile()
        ElseIf Main.ComboBox1.SelectedItem = "OBO Betterman" Then
            OpenOBOBettermanInvoiceFile()
        End If
    End Sub

    Public Sub UploadInvoiceFile()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Main.TextBox1.Text = "3046" Then
            UploadABBInvoiceFile()
        ElseIf Main.TextBox1.Text = "1029" Then
            UploadOBOBettermanInvoiceFile()
        End If
    End Sub

    Public Function LoadInvoiceFromTMPTable() As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� � �������� ABB � Scala �� ��������� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String
        Dim MySQLStr As String
        Dim MyRec1 As ADODB.Recordset
        Dim cmd As New ADODB.Command                    '������� (spp ���������)
        Dim MyParam As ADODB.Parameter                  '������������ �������� ����� 1
        Dim MyParam1 As ADODB.Parameter                 '������������ �������� ����� 2
        Dim MyParam2 As ADODB.Parameter                 '������������ �������� ����� 3
        Dim MyParam3 As ADODB.Parameter                 '������������ �������� ����� 4
        Dim MyParam4 As ADODB.Parameter                 '������������ �������� ����� 5
        Dim MyParam5 As ADODB.Parameter                 '������������ �������� ����� 6
        Dim MyParam6 As ADODB.Parameter                 '������������ �������� ����� 7
        Dim MyParam7 As ADODB.Parameter                 '������������ �������� ����� 8
        Dim MyParam8 As ADODB.Parameter                 '������������ �������� ����� 9
        Dim MyParam9 As ADODB.Parameter                 '������������ �������� ����� 10
        Dim MyParam10 As ADODB.Parameter                '������������ �������� ����� 11 
        Dim MyParam11 As ADODB.Parameter                '������������ �������� ����� 12
        Dim MyParam12 As ADODB.Parameter                '������������ �������� ����� 13
        Dim MyParam13 As ADODB.Parameter                '������������ �������� ����� 14 ���� ������������
        Dim MyParam14 As ADODB.Parameter                '������������ �������� ����� 15 ���� ������������

        MyRezStr = ""
        InitMyConn(False)
        Try
            cmd.ActiveConnection = declarations.MyConn
            cmd.CommandText = "spp_PurchaseInvoice_AutoUploadLine_R2"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.CommandTimeout = 300
            MyRec1 = New ADODB.Recordset

            '----�������� ����������---------------------------------------------------
            '---����� ������ �� �������
            MyParam = cmd.CreateParameter("@ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam)
            '---����� ���� �������
            MyParam1 = cmd.CreateParameter("@Invoice", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam1)
            '---���� ���� �������
            MyParam2 = cmd.CreateParameter("@InvoiceDateSTR", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam2)
            '---��� ������ ���� �������
            MyParam3 = cmd.CreateParameter("@CurrCode", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam3)
            '---��� ��������
            MyParam4 = cmd.CreateParameter("@MySalesmanCode", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam4)
            '---��� ��������
            MyParam5 = cmd.CreateParameter("@MySalesmanName", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam5)
            '---���� ������ �� ����������� �������
            MyParam6 = cmd.CreateParameter("@PurchInvoiceExRate", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam6)
            '---����� ������������������ ������ �� �������
            MyParam7 = cmd.CreateParameter("@POrder", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam7)
            '---��� ������ ����������
            MyParam8 = cmd.CreateParameter("@ItemSuppCode", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 50)
            cmd.Parameters.Append(MyParam8)
            '---����������
            MyParam9 = cmd.CreateParameter("@QTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam9)
            '---����� ��� ��� �� ������
            MyParam10 = cmd.CreateParameter("@Price", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam10)
            '---������
            MyParam11 = cmd.CreateParameter("@Country", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam11)
            '---���
            MyParam12 = cmd.CreateParameter("@GTD", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 255)
            cmd.Parameters.Append(MyParam12)
            '---������������ �������� (������)
            MyParam13 = cmd.CreateParameter("@MyRezStr", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
            cmd.Parameters.Append(MyParam13)
            '---������������ �������� (Double)
            MyParam14 = cmd.CreateParameter("@MyRestQTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamOutput)
            cmd.Parameters.Append(MyParam14)


            MySQLStr = "SELECT #_MyInvoice.* "
            MySQLStr = MySQLStr & "FROM #_MyInvoice WITH(NOLOCK) "
            MySQLStr = MySQLStr & "ORDER BY SupplierItemCode, QTY DESC "
            InitMyConn(False)
            MyRec1.Open(MySQLStr, declarations.MyConn)
            If MyRec1.EOF = True And MyRec1.BOF = True Then
                trycloseMyRec()
            Else
                MyRec1.MoveFirst()
                While MyRec1.EOF = False
                    '----�������� ����������---------------------------------------------------
                    '---ID ������
                    MyParam.Value = MyRec1.Fields("ID").Value
                    '---����� ���� �������
                    MyParam1.Value = MyRec1.Fields("Invoice").Value
                    '---���� ���� �������
                    MyParam2.Value = Format(MyRec1.Fields("InvoiceDate").Value, "dd/MM/yyyy")
                    '---��� ������ ���� �������
                    MyParam3.Value = MyRec1.Fields("InvoiceCurrCode").Value
                    '---��� ��������
                    MyParam4.Value = MyRec1.Fields("SalesmanCode").Value
                    '---��� ��������
                    MyParam5.Value = MyRec1.Fields("SalesmanName").Value
                    '---���� ������ �� ����������� �������
                    MyParam6.Value = MyRec1.Fields("InvoiceCurrExchRate").Value
                    '---����� ������ �� �������
                    MyParam7.Value = MyRec1.Fields("ConsPurchaseOrderNum").Value
                    '---��� ������ ����������
                    MyParam8.Value = MyRec1.Fields("SupplierItemCode").Value
                    '---����������
                    MyParam9.Value = MyRec1.Fields("RestQTY").Value
                    '---����� ��� ��� �� ������
                    MyParam10.Value = MyRec1.Fields("SummWithoutVAT").Value
                    '---������
                    MyParam11.Value = MyRec1.Fields("Country").Value
                    '---���
                    MyParam12.Value = MyRec1.Fields("GTD").Value

                    '---������ �������� ���������------------------------------------------------
                    '---����������
                    '--SetBlock(MyParam.Value, 1)   --��������� � �������� ���������

                    '---���������
                    cmd.Execute()
                    MyRezStr = MyRezStr + LTrim(RTrim(MyParam13.Value))

                    '---������ ����������
                    '--RemoveBlock()

                    '---������������� ��������
                    MySQLStr = "UPDATE #_MyInvoice "
                    MySQLStr = MySQLStr & "SET RestQTY = " & Replace(CStr(MyParam14.Value), ".", ",") & " "
                    MySQLStr = MySQLStr & "WHERE (ID = " & CStr(MyParam.Value) & ") "
                    InitMyConn(False)
                    Try
                        declarations.MyConn.Execute(MySQLStr)
                    Catch
                        MyRezStr = MyRezStr + " ������ ������������� �������� ��� ������ ���������� " + MyParam8.Value + " �������� ���������� " + MyParam9.Value + " ��������� ������� ����� ������." + Chr(13)
                    End Try
                    MyRec1.MoveNext()
                    Main.progressBar1.Increment(1)
                End While
                MyRec1.Close()
                MyRec1 = Nothing
            End If

            LoadInvoiceFromTMPTable = MyRezStr
        Catch ex As Exception
            If MyRec Is Nothing Then
            Else
                If MyRec1.State <> 0 Then
                    MyRec1.Close()
                End If
            End If
            MyRezStr = MyRezStr + ex.Message
            LoadInvoiceFromTMPTable = MyRezStr
        End Try
    End Function

    Public Function UploadingRezult(ByVal MyRezStr As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ����������� �������� � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MyErrorForm = New ErrorForm

        If MyRezStr = "" Then
        Else
            MyErrorForm.MyHdr = "�� ����� ������� �� ���� ������ " & Chr(13)
        End If

        '------------����� ���������� � ������� �� �������, �� ������� ���� �������
        If MyRezStr <> "" Then
            MyRezStr = MyRezStr & Chr(13)
        End If

        MySQLStr = "SELECT PC190300.PC19001 AS OrderNum, "
        MySQLStr = MySQLStr & "PC010300.PC01023 AS WhNum "
        MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC190300.PC19001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC190300.PC19012 = N'" & Main.textBox3.Text & "') AND "
        MySQLStr = MySQLStr & "(PC190300.PC19010 = dateadd( day, datediff(day, 0, GETDATE()), 0)) "
        MySQLStr = MySQLStr & "GROUP BY PC190300.PC19001, "
        MySQLStr = MySQLStr & "PC010300.PC01023 "
        MySQLStr = MySQLStr & "ORDER BY OrderNum "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MyRezStr = MyRezStr & "������ �� N " & Main.textBox3.Text & " �� ���������� �� ��� ������ ������ �� ������� " & Chr(13)
        Else
            MyRezStr = MyRezStr & "������ �� N " & Main.textBox3.Text & " ���������� ��� ��������� ������� �� �������: " & Chr(13)
            MyRezStr = MyRezStr & "����� �� �������    ����� ������" & Chr(13)
            declarations.MyRec.MoveFirst()
            While declarations.MyRec.EOF = False
                MyRezStr = MyRezStr & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("OrderNum").Value & "                    ", 22) & declarations.MyRec.Fields("WhNum").Value & Chr(13)
                declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        '------------����� ���������� � �������������
        MyRezStr = MyRezStr & Chr(13)

        MySQLStr = "SELECT SupplierItemCode, QTY, RestQTY "
        MySQLStr = MySQLStr & "FROM #_MyInvoice WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (RestQTY <> 0) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            MyRezStr = MyRezStr & "� ���������� ������� �� N " & Main.textBox3.Text & " �� ���� ������� ��������� ������: " & Chr(13)
            MyRezStr = MyRezStr & "��� ������ ����������               ���������� � ��     ���������� ����������" & Chr(13)
            declarations.MyRec.MoveFirst()
            While declarations.MyRec.EOF = False
                MyRezStr = MyRezStr & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("SupplierItemCode").Value & "                                    ", 36) & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("QTY").Value.ToString & "                    ", 20) & MyRec.Fields("RestQTY").Value.ToString & Chr(13)
                declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If


        MyErrorForm.MyMsg = MyRezStr & Chr(13)
        MyErrorForm.ShowDialog()

        Main.button3.Enabled = False
    End Function

    Public Function CheckUOMInOrders() As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��� ������� ��������� � ������ ��������� � ��������� ��������� �
        '// �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRezStr As String

        MyRezStr = ""

        MySQLStr = "SELECT PC030300.PC03001 AS OrderN, PC030300.PC03005 AS ItemN, View_1_1.txt AS OrderUOM, View_1.txt AS CardUOM "
        MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT ConsPurchaseOrderNum "
        MySQLStr = MySQLStr & "FROM #_MyInvoice WITH(NOLOCK) "
        MySQLStr = MySQLStr & "GROUP BY ConsPurchaseOrderNum) AS View_1_2 ON "
        MySQLStr = MySQLStr & "PC010300.PC01052 = View_1_2.ConsPurchaseOrderNum INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 AND PC030300.PC03009 <> SC010300.SC01134 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1 ON SC010300.SC01134 = View_1.num INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS num, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE   (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')) AS View_1_1 ON PC030300.PC03009 = View_1_1.num "
        MySQLStr = MySQLStr & "ORDER BY OrderN, ItemN "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            MyRezStr = MyRezStr & "� ���������� ������ ���� ������, ������� ��������� ������� ���������� �� ������� ���������, ��������� " & Chr(13) & "� �������� ������. " & Chr(13)
            MyRezStr = MyRezStr & "����� �� �������    ��� ������                          ��. ��������� � �-�� ��. ��������� � ��������" & Chr(13)
            declarations.MyRec.MoveFirst()
            While declarations.MyRec.EOF = False
                MyRezStr = MyRezStr & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("OrderN").Value & "                    ", 20) & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("ItemN").Value & "                                    ", 36) & Microsoft.VisualBasic.Left(declarations.MyRec.Fields("OrderUOM").Value & "                     ", 21) & declarations.MyRec.Fields("CardUOM").Value & Chr(13)
                declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        CheckUOMInOrders = MyRezStr

    End Function

    Public Function CheckEmptyInOrders() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '// 
        '// �������� ��� �� ��������� �������� ��� ����� � ������� �������� (����� ���)
        '// 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM #_MyInvoice "
        MySQLStr = MySQLStr & "WHERE (Ltrim(Rtrim(SupplierItemCode)) = '') "
        MySQLStr = MySQLStr & "OR (QTY = 0) "
        MySQLStr = MySQLStr & "OR (SummWithoutVAT = 0) "
        MySQLStr = MySQLStr & "OR (Ltrim(Rtrim(Country)) = '') "
        InitMyRec(False, MySQLStr)
        If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            trycloseMyRec()
            CheckEmptyInOrders = False
        Else
            If declarations.MyRec.Fields("CC").Value = 0 Then
                CheckEmptyInOrders = True
            Else
                CheckEmptyInOrders = False
            End If
            trycloseMyRec()
        End If
    End Function
End Module
