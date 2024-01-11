Module ShippingCheckFunctions
    Public Function ExecShippingAllovance(ByVal MyOrderID As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ���������� �� �������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim CustomerType As String
        Dim MyRez As Object
        Dim MySQLStr As String

        '---���� �������� � ���������
        MySQLStr = "Update OR010300 "
        MySQLStr = MySQLStr & "Set OR01016 = View_1.CC "
        MySQLStr = MySQLStr & "FROM OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
        MySQLStr = MySQLStr & "FROM OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrderID) & "') "
        MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON OR010300.OR01001 = View_1.OR03001 "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        If CheckBlock(MyOrderID) = True Then
            CustomerType = CheckCustomerType(MyOrderID)
            If CustomerType = "����������" Then
            Else
                '--============�������� - �������� �� ����� ��������� ==================================================
                MySQLStr = "SELECT IsProject, CASE WHEN ProjectID IS NULL THEN '' ELSE CONVERT(nvarchar(40), ProjectID) END AS ProjectID "
                MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo "
                MySQLStr = MySQLStr & "WHERE (OrderID = N'" & MyOrderID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Declarations.MyProjectIsApproved = 0
                    MyProjectID = ""
                    trycloseMyRec()
                Else
                    Declarations.MyProjectIsApproved = Declarations.MyRec.Fields("IsProject").Value
                    Declarations.MyProjectID = Declarations.MyRec.Fields("ProjectID").Value
                    trycloseMyRec()
                End If

                '---�������� - ���� �� ������ � ������������ �������� � CRM
                If Declarations.MyProjectIsApproved <> 0 Then    '----------------��������� �����
                    MySQLStr = "SELECT COUNT(tbl_CRM_Projects.ProjectID) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_CRM_Projects ON tbl_SalesHdr_ProjectAddInfo.ProjectID = tbl_CRM_Projects.ProjectID "
                    MySQLStr = MySQLStr & "WHERE (tbl_SalesHdr_ProjectAddInfo.OrderID = N'" & MyOrderID & "') "
                    MySQLStr = MySQLStr & "AND (tbl_CRM_Projects.IsApproved = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("���������� ���������, �������� �� � CRM ���������� �� �������. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
                        Exit Function
                    Else
                        If Declarations.MyRec.Fields("CC").Value <= 0 Then
                            MsgBox("������������� �������, ���������� ��� ������� ������, � CRM ���. ����� ����� �� ����� ���� ��������. ������� �������� ���������� � CRM � ��������� �� � ��������� �� ��������.", vbCritical, "��������!")
                            trycloseMyRec()
                            Exit Function
                        Else
                            trycloseMyRec()
                        End If
                    End If
                End If

                '---�������� - ��������� �� ������������ ������� � CRM
                If Declarations.MyProjectIsApproved <> 0 Then    '----------------��������� �����
                    MySQLStr = "SELECT COUNT(tbl_CRM_Project_Details.ProjectID) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_CRM_Project_Details ON tbl_SalesHdr_ProjectAddInfo.ProjectID = tbl_CRM_Project_Details.ProjectID "
                    MySQLStr = MySQLStr & "WHERE (tbl_SalesHdr_ProjectAddInfo.OrderID = N'" & MyOrderID & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("���������� ���������, ��������� �� ������������ �� �������. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
                        Exit Function
                    Else
                        If Declarations.MyRec.Fields("CC").Value <= 0 Then
                            MsgBox("�� �������, ���������� ��� ������� ������, ���������� �� �������� �� ���������� ������������. ����� ����� �� ����� ���� ��������.", vbCritical, "��������!")
                            trycloseMyRec()
                            Exit Function
                        Else
                            trycloseMyRec()
                        End If
                    End If
                End If

                '---�������� - ��� ������ � ������ � ������ � ������� CRM - ���� � ��� ��
                If Declarations.MyProjectIsApproved <> 0 Then    '----------------��������� �����
                    MySQLStr = "SELECT tbl_CRM_Companies.ScalaCustomerCode, OR010300.OR01003 "
                    MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_CRM_Projects ON tbl_SalesHdr_ProjectAddInfo.ProjectID = tbl_CRM_Projects.ProjectID INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
                    MySQLStr = MySQLStr & "OR010300 ON tbl_SalesHdr_ProjectAddInfo.OrderID = OR010300.OR01001 "
                    MySQLStr = MySQLStr & "WHERE (tbl_SalesHdr_ProjectAddInfo.OrderID = N'" & MyOrderID & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        MsgBox("���������� ������� ��� ���������� � Scala � ����� ����������, ��������� ��� ���������������� �������. ��������, ����� - �� �� ������� �������. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
                        Exit Function
                    Else
                        If Trim(Declarations.MyRec.Fields("ScalaCustomerCode").Value) <> Trim(Declarations.MyRec.Fields("OR01003").Value) Then
                            MsgBox("��� ���������� � Scala: " & Trim(Declarations.MyRec.Fields("OR01003").Value) & " �� ��������� � ����� ����������, ��������� ��� ���������������� �������: " & Trim(Declarations.MyRec.Fields("ScalaCustomerCode").Value) & ". ��������������� ����������, ����� ���� ��������� ������� ������ ���������� �� ��������.", vbCritical, "��������!")
                            trycloseMyRec()
                            Exit Function
                        Else
                            trycloseMyRec()
                        End If
                    End If
                End If

                '---==============�������� ������ �����=========================================================
                Declarations.IsWEBOrder = CheckWEBOrNot(MyOrderID)

                If Declarations.IsWEBOrder = 0 Then '---�� �������� ������� � WEB ����� - ��������� �����
                    If Declarations.MyProjectIsApproved = 0 Then  '----------------����������� �����
                        If CheckNegativeMargin(MyOrderID, 1) = False Then
                            'If CheckRights1(Declarations.UserCode, "ShipmentsWithLowMarginLevel1", Trim(MyOrderID)) = "���������" Then
                            If CheckRights1(Declarations.UserCode, "ShipmentsWithLowMarginLevel2", Trim(MyOrderID)) = "���������" Then
                                '============�������� ��� �������� ����������!!!
                                MsgBox("� ������ ���� ������ � ������ ���� �������������. ��������� ����� ����� ����� ������ ����������, ���������� ����� ������.", vbOKOnly, "��������!")
                                Exit Function
                            Else
                                MyRez = MsgBox("� ������ ���� ������ � ������ ���� �������������. ���������� �������� ������ ������?", vbYesNo, "��������!")
                                If MyRez = vbNo Then
                                    Exit Function
                                Else
                                    '---���������� ������
                                    SaveMarginInfo("ShipmentsWithLowMarginLevel2", MyOrderID, Declarations.UserCode)
                                End If
                            End If
                            'Else
                            '---������ ��������� ������� � 1 ���, ���� ����� ������������� ShipmentsWithLowMarginLevel1
                            '---���� ����� ��������� - ����������������� ��, ��� ����.
                            'MyRez = MsgBox("� ������ ���� ������ � ������ ���� �������������. ���������� �������� ������ ������?", vbYesNo, "��������!")
                            'If MyRez = vbNo Then
                            'Exit Function
                            'Else
                            '---���������� ������
                            'SaveMarginInfo("ShipmentsWithLowMarginLevel1", MyOrderID, Declarations.UserCode)
                            'End If
                            'End If
                        End If
                    Else                                        '----------------��������� �����
                        If CheckNegativeProjectMargin(MyOrderID) = False Then
                            'If CheckRights1(Declarations.UserCode, "ShipmentsWithLowMarginLevel1", Trim(MyOrderID)) = "���������" Then
                            If CheckRights1(Declarations.UserCode, "ShipmentsWithLowMarginLevel2", Trim(MyOrderID)) = "���������" Then
                                '============�������� ��� �������� ����������!!!
                                MsgBox("� ������ ���� ������ � ������ ���� �������������. ��������� ����� ����� ����� ������ ����������, ���������� ����� ������.", vbOKOnly, "��������!")
                                Exit Function
                            Else
                                MyRez = MsgBox("� ������ ���� ������ � ������ ���� �������������. ���������� �������� ������ ������?", vbYesNo, "��������!")
                                If MyRez = vbNo Then
                                    Exit Function
                                Else
                                    '---���������� ������
                                    SaveMarginInfo("ShipmentsWithLowMarginLevel2", MyOrderID, Declarations.UserCode)
                                End If
                            End If
                            'Else
                            '---������ ��������� ������� � 1 ���, ���� ����� ������������� ShipmentsWithLowMarginLevel1
                            '---���� ����� ��������� - ����������������� ��, ��� ����.
                            'MyRez = MsgBox("� ������ ���� ������ � ������ ���� �������������. ���������� �������� ������ ������?", vbYesNo, "��������!")
                            'If MyRez = vbNo Then
                            'Exit Function
                            'Else
                            '---���������� ������
                            'SaveMarginInfo("ShipmentsWithLowMarginLevel1", MyOrderID, Declarations.UserCode)
                            'End If
                            'End If
                        End If
                    End If
                End If

                '---=============�������� ����������� ��������� ��� ����� �������� � 1 ���=========================
                If CheckSalesPrices(MyOrderID) = False Then
                    If CheckRights(Declarations.UserCode, "ShipmentsOverLimit") = "���������" Then
                        MsgBox("����� �������� ������ � 1 ��� ���������� ����. ��������� ����� ����� ����� ������ ����������, ���������� ����� ������.", vbOKOnly, "��������!")
                        Exit Function
                    Else
                        MyRez = MsgBox("����� �������� ������ � 1 ��� ���������� ����. ���������� �������� ������ ������?", vbYesNo, "��������!")
                        If MyRez = vbNo Then
                            Exit Function
                        Else
                        End If
                    End If
                End If
                End If

                If CheckShippingAllowed(MyOrderID) = True Then
                    '---���������� �������� 3 (��������� ��������) � ������
                    MySQLStr = "UPDATE OR010300 "
                    MySQLStr = MySQLStr & "SET OR01008 = 3 "
                    MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyOrderID & "')"
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---��������� ���������� � ����� ������ ���������� �� ��������
                    MySQLStr = "INSERT INTO tbl_ShipmentsAllowanceInfo "
                    MySQLStr = MySQLStr & "(ID, OrderID, AllowanceData, SalesmanCode, SalesmanName) "
                    MySQLStr = MySQLStr & "VALUES ("
                    MySQLStr = MySQLStr & "NEWID(), "
                    MySQLStr = MySQLStr & "N'" & MyOrderID & "', "
                    MySQLStr = MySQLStr & "dateadd( day, datediff(day, 0, GETDATE()), 0), "
                    MySQLStr = MySQLStr & "N'" & Declarations.SalesmanCode & "', "
                    MySQLStr = MySQLStr & "N'" & Declarations.SalesmanName & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---��������� ��������� ���������� � ������ �� ������ ������ ���������� �� ��������
                    MySQLStr = "INSERT INTO tbl_ShipmentsAllowanceDetailInfo "
                    MySQLStr = MySQLStr & "SELECT NEWID(), OR030300.OR03001, ISNULL(View_17.Expr1, N'') + ISNULL(View_16.ScalaComment, N'') AS Comments, DATEADD(day, DATEDIFF(day, 0, GETDATE()), "
                    MySQLStr = MySQLStr & "0) AS AllowanceDate, OR030300.OR03002, OR030300.OR03005 AS ItemCode, LTRIM(RTRIM(LTRIM(RTRIM(OR030300.OR03006)) "
                    MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(OR030300.OR03007)))) AS ItemName, OR030300.OR03011 AS Ordered, OR030300.OR03053 AS Reserved, "
                    MySQLStr = MySQLStr & "OR030300.OR03010 AS UOM "
                    MySQLStr = MySQLStr & "FROM OR030300 LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "(SELECT OR17001 AS OrderID, LTRIM(RTRIM(LTRIM(RTRIM(OR17005)) + ' ' + LTRIM(RTRIM(OR17006)))) AS ScalaComment "
                    MySQLStr = MySQLStr & "FROM OR170300 "
                    MySQLStr = MySQLStr & "WHERE (OR17002 = N'000000') AND (OR17004 = N'510')) AS View_16 ON OR030300.OR03001 = View_16.OrderID LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "(SELECT OrderID, CASE WHEN ISNULL(CustomerPONum, '') = '' THEN '' ELSE ' N ������ �������: ' + LTRIM(Rtrim(ISNULL(CustomerPONum, ''))) "
                    MySQLStr = MySQLStr & "END + CASE WHEN ISNULL(CustomerAgreementNum, '') = '' THEN '' ELSE ' N ��������: ' + Ltrim(Rtrim(ISNULL(CustomerAgreementNum, "
                    MySQLStr = MySQLStr & "''))) END + CASE WHEN ISNULL(CustomerManagerName, '') "
                    MySQLStr = MySQLStr & "= '' THEN '' ELSE ' �������� (�������): ' + Ltrim(Rtrim(ISNULL(CustomerManagerName, ''))) "
                    MySQLStr = MySQLStr & "END + CASE WHEN ISNULL(DeliveryAddress, '') = '' THEN '' ELSE ' ��������: ' + Ltrim(Rtrim(ISNULL(DeliveryAddress, ''))) "
                    MySQLStr = MySQLStr & "END AS Expr1 "
                    MySQLStr = MySQLStr & "FROM tbl_SalesHdr_EDOInfo) AS View_17 ON OR030300.OR03001 = View_17.OrderID "
                    MySQLStr = MySQLStr & "WHERE (OR030300.OR03001 = N'" & MyOrderID & "') "
                    MySQLStr = MySQLStr & "AND (OR030300.OR03003 = N'000000') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---� �������� ������ ��������
                    'GetShippingList(MyOrderID)
                Else
                End If
        End If
    End Function

    Public Function CheckBlock(ByVal MyOrderID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ������� - ������������ ��� ���
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustomerID As String                    '��� �������
        Dim CustomerIsCredit As Integer             '�������� ��������� (1) ��� ��� (0)
        Dim CustomerIsBlocked As Integer            '������������ (1) ��� ��� (0)

        MySQLStr = "SELECT OR01003 "
        MySQLStr = MySQLStr & "FROM OR010300 "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyOrderID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
            MsgBox("� ������ " & MyOrderID & " ���������� ���������� ��� ����������. ��������, ����� ������. ���������� � ��������������. ", vbCritical, "��������!")
            CheckBlock = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.CustomerID = Declarations.MyRec.Fields("OR01003").Value
            CustomerID = Declarations.CustomerID
            trycloseMyRec()
        End If

        CustomerIsCredit = 0
        MySQLStr = "SELECT tbl_CustomerCard0300.IsBlocked, CASE WHEN (SL01024 = N'0' OR "
        MySQLStr = MySQLStr & "SL01024 = N'00') AND SL01037 = 0 THEN 0 ELSE 1 END AS IsCredit "
        MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 INNER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON tbl_CustomerCard0300.SL01001 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_CustomerCard0300.SL01001 = N'" & Declarations.CustomerID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
            MsgBox("��� ���������� " & CustomerID & " ���������� ���������� ��������� �� ��� ���. ���������� � ��������������. ", vbCritical, "��������!")
            CheckBlock = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            CustomerIsBlocked = Declarations.MyRec.Fields("IsBlocked").Value
            If Declarations.MyRec.Fields("IsCredit").Value = 1 Then     '---���������
                CustomerIsCredit = 1
                trycloseMyRec()
                If CustomerIsBlocked = 1 Then
                    '---������ ��������� � ������������ - ��� ������������� � ��������� �������� �������
                    MsgBox("���������� " & CustomerID & " �������� ��������� � ������������. ��� ������������� ������� ��������� ��������� ���������� ���������� �������� �� ������� (������ ��� ����������� ������������). ", vbCritical, "��������!")
                    CheckBlock = False
                    Exit Function
                End If
            End If
        End If

        '---�������� - �� ����� �� ���� �������� (��� ��������� �����������)-------------------
        If CustomerIsCredit = 1 Then                    '---��������� ������ ��� ��������� �����������
            MySQLStr = "SELECT DataFrom, DataTo "
            MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 "
            MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Declarations.CustomerID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("������ �������� - �� ����� �� ���� ��������. ���������� � ��������������.", vbCritical, "��������!")
                CheckBlock = False
                trycloseMyRec()
                Exit Function
            Else
                Declarations.MyRec.MoveFirst()
                If Declarations.MyRec.Fields("DataFrom").Value = CDate("01/01/1900") Or Declarations.MyRec.Fields("DataTo").Value = CDate("01/01/1900") Then
                    '---���� �� ����������� - �� ���������
                    trycloseMyRec()
                Else
                    If Declarations.MyRec.Fields("DataFrom").Value <= Now() And Declarations.MyRec.Fields("DataTo").Value > Now() Then
                        '---��� OK, �� �������� ������� ��������� �� ��������� ��������
                        If DateDiff("d", Now(), Declarations.MyRec.Fields("DataTo").Value) < 60 Then
                            MsgBox("�� ����� �������� �������� � ����������� �������� ������ ���� �������. ������� ���� � ���������� ������ �������� � �������� ����� ������ � ����.", vbOKOnly, "��������!")
                        End If
                        trycloseMyRec()
                    Else
                        MsgBox("����� ������ ���������, ��� ��� ���������� ��� ��� �� ������� ���� �������� �������� �������� � ������ ��������. ��������� ����� ������� � �������� ������ � ��� � ����.", vbCritical, "��������!")
                        CheckBlock = False
                        trycloseMyRec()
                        Exit Function
                    End If
                End If
            End If
        End If

        CheckBlock = True
    End Function

    Public Function CheckCustomerType(ByVal OrderNumber As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������� - ��������� ��� ���
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim AddStr As String
        Dim MyRez As Object

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT dbo.SL010300.SL01001 AS CustomerCode, "
        MySQLStr = MySQLStr & "dbo.SL010300.SL01002 AS CustomerName, "
        MySQLStr = MySQLStr & "RTRIM(LTRIM(dbo.SL010300.SL01024)) AS Terms, "
        MySQLStr = MySQLStr & "RTRIM(LTRIM(dbo.SL010300.SL01037)) AS Amount "
        MySQLStr = MySQLStr & "FROM dbo.OR010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.SL010300 ON dbo.OR010300.OR01003 = dbo.SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR010300.OR01001 = N'" & Right("0000000000" & OrderNumber, 10) & "') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT DISTINCT dbo.SL010300.SL01001 AS CustomerCode, "
        MySQLStr = MySQLStr & "dbo.SL010300.SL01002 AS CustomerName, "
        MySQLStr = MySQLStr & "RTRIM(LTRIM(dbo.SL010300.SL01024)) AS Terms, "
        MySQLStr = MySQLStr & "RTRIM(LTRIM(dbo.SL010300.SL01037)) AS amount "
        MySQLStr = MySQLStr & "FROM dbo.OR200300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.SL010300 ON dbo.OR200300.OR20003 = dbo.SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR200300.OR20001 = N'" & Right("0000000000" & OrderNumber, 10) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckCustomerType = "�����������"
            trycloseMyRec()
            Exit Function
        End If
        Declarations.MyRec.MoveFirst()
        Declarations.CustomerID = Trim(Declarations.MyRec.Fields("CustomerCode").Value)
        '---�����. �����������
        MyRez = InStr(UCase(Declarations.MyRec.Fields("CustomerCode").Value), UCase("INTR"))
        If MyRez <> 0 Then
            CheckCustomerType = "����������"
            trycloseMyRec()
            Exit Function
        End If
        '---��������� ���������� (�� ���������� � ����������)
        MyRez = InStr(UCase(Declarations.MyRec.Fields("CustomerName").Value), UCase("��������� ����������"))
        If MyRez <> 0 Then
            CheckCustomerType = "���������"
            trycloseMyRec()
            Exit Function
        End If
        '---��������� - ��������� � �����������
        AddStr = Declarations.MyRec.Fields("Terms").Value
        Declarations.CreditAmount = CDbl(Declarations.MyRec.Fields("Amount").Value)
        trycloseMyRec()
        MySQLStr = "SELECT SL22005 "
        MySQLStr = MySQLStr & "FROM SL220300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL22002 = N'" & AddStr & "') "
        MySQLStr = MySQLStr & "AND (SL22001 = N'0')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.CreditInDays = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.CreditInDays = CInt(Declarations.MyRec.Fields("SL22005").Value)
        End If
        If (AddStr = "0" Or AddStr = "00") And Declarations.CreditAmount = 0 Then
            CheckCustomerType = "�����������"
        Else
            CheckCustomerType = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 4")
        CheckCustomerType = "���������������������"
    End Function

    Public Function CheckNegativeMargin(ByVal OrderNum As String, ByVal MarginType As Integer) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ������
        '// MarginType: 0 - �� ��������� ������������� 1 - �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            Return CheckNegativeMarginLO(OrderNum, MarginType)
        Else
            Return CheckNegativeMarginExcel(OrderNum, MarginType)
        End If
    End Function

    Public Function CheckNegativeMarginExcel(ByVal OrderNum As String, ByVal MarginType As Integer) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ������
        '// MarginType: 0 - �� ��������� ������������� 1 - �� ����������� ����� � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer
        'Dim MyScala As New SfwIII.Application

        On Error GoTo MyCatch

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)

        If MarginType = 0 Then  '---����� �� ��������� �������������
            MySQLStr = "EXEC spp_ShippingCheck_GetNegativeMarginInfo "
        Else                    '---����� �� ����������� �������������
            MySQLStr = "EXEC spp_ShippingCheck_GetRNegMarginInfo "
        End If
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"

        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CheckNegativeMarginExcel = True
        Else
            '---���� ������������� �����, �� � �������
            MyObj = CreateObject("Excel.Application")
            MyObj.SheetsInNewWorkbook = 1
            MyWRKBook = MyObj.Workbooks.Add

            MyWRKBook.ActiveSheet.Columns("A:E").ColumnWidth = 15
            MyWRKBook.ActiveSheet.Columns("C").ColumnWidth = 40

            '---��������� �����
            MyWRKBook.ActiveSheet.Range("B1") = "����� �� ������� ����� " & OrderNum
            MyWRKBook.ActiveSheet.Range("B2") = "� ���� ������ ���������� ������ � �������� ������ ���� ����������� "
            MyWRKBook.ActiveSheet.Range("B3") = "��� ��������� ����������� �� ��������  "
            MyWRKBook.ActiveSheet.Range("B4") = "���������� ���������� � ��������� �� ��������.  "
            MyWRKBook.ActiveSheet.Range("A6") = "������ ������� � ����������� ������ ���� �����������:"

            MyWRKBook.ActiveSheet.Range("B1").Select()
            MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
            MyWRKBook.ActiveSheet.Range("A6").Select()
            MyWRKBook.ActiveSheet.Range("A6").Font.Bold = True
            '---��������� �������� ������
            MyWRKBook.ActiveSheet.Range("A7") = "N ������"
            MyWRKBook.ActiveSheet.Range("B7") = "��� ��������"
            MyWRKBook.ActiveSheet.Range("C7") = "��� ��������"
            MyWRKBook.ActiveSheet.Range("D7") = "����"
            MyWRKBook.ActiveSheet.Range("A7:D7").Select()
            MyWRKBook.ActiveSheet.Range("A7:D7").Borders(5).LineStyle = -4142
            MyWRKBook.ActiveSheet.Range("A7:D7").Borders(6).LineStyle = -4142
            With MyWRKBook.ActiveSheet.Range("A7:D7").Borders(7)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A7:D7").Borders(8)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A7:D7").Borders(9)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A7:D7").Borders(10)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A7:D7").Borders(11)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A7:D7").Interior
                .ColorIndex = 36
                .Pattern = 1
                .PatternColorIndex = -4105
            End With

            '---����� ����� ������
            i = 8
            Declarations.MyRec.MoveFirst()
            Declarations.MinMargin = Declarations.MyRec.Fields("MinMargin").Value
            While Not Declarations.MyRec.EOF
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)) = Declarations.MyRec.Fields("Str").Value
                MyWRKBook.ActiveSheet.Range("B" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" + CStr(i)) = Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("C" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" + CStr(i)) = Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("D" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" + CStr(i)) = Declarations.MyRec.Fields("COST").Value

                Declarations.MyRec.MoveNext()
                i = i + 1
            End While

            MyWRKBook.ActiveSheet.Range("A1").Select()
            MyObj.Application.Visible = True
            MyObj = Nothing
            CheckNegativeMarginExcel = False
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 10")
    End Function

    Public Function CheckNegativeMarginLO(ByVal OrderNum As String, ByVal MarginType As Integer) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ������
        '// MarginType: 0 - �� ��������� ������������� 1 - �� ����������� ����� � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim i As Integer

        On Error GoTo MyCatch

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)
        If MarginType = 0 Then  '---����� �� ��������� �������������
            MySQLStr = "EXEC spp_ShippingCheck_GetNegativeMarginInfo "
        Else                    '---����� �� ����������� �������������
            MySQLStr = "EXEC spp_ShippingCheck_GetRNegMarginInfo "
        End If
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CheckNegativeMarginLO = True
        Else
            '---���� ������������� �����, �� � �������
            LOSetNotation(0)
            oServiceManager = CreateObject("com.sun.star.ServiceManager")
            oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
            oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
            Dim arg(1)
            arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
            arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
            oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
            oSheet = oWorkBook.getSheets().getByIndex(0)
            oFrame = oWorkBook.getCurrentController.getFrame
            '---������ �������
            oSheet.getColumns().getByName("A").Width = 2000
            oSheet.getColumns().getByName("B").Width = 5000
            oSheet.getColumns().getByName("C").Width = 10000
            oSheet.getColumns().getByName("D").Width = 4000
            '---��������� �����
            oSheet.getCellRangeByName("B1").String = "����� �� ������� ����� " & OrderNum
            oSheet.getCellRangeByName("B2").String = "� ���� ������ ���������� ������ � �������� ������ ���� �����������"
            oSheet.getCellRangeByName("B3").String = "��� ��������� ����������� �� ��������"
            oSheet.getCellRangeByName("B4").String = "���������� ���������� � ��������� �� ��������."
            oSheet.getCellRangeByName("A6").String = "������ ������� � ����������� ������ ���� �����������:"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:B6", "Tahoma")
            LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1:B1")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1:B6", 11)
            '---��������� �������
            oSheet.getCellRangeByName("A7").String = "N ������"
            oSheet.getCellRangeByName("B7").String = "��� ������"
            oSheet.getCellRangeByName("C7").String = "��� ��������"
            oSheet.getCellRangeByName("D7").String = "����"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A7:D7", "Calibri")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A7:D7", 9)
            LOWrapText(oServiceManager, oDispatcher, oFrame, "A7:D7")
            oSheet.getCellRangeByName("A7:D7").CellBackColor = RGB(114, 251, 238)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
            Dim LineFormat As Object
            LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
            LineFormat.LineStyle = 0
            LineFormat.LineWidth = 70
            oSheet.getCellRangeByName("A7:D7").TopBorder = LineFormat
            oSheet.getCellRangeByName("A7:D7").RightBorder = LineFormat
            oSheet.getCellRangeByName("A7:D7").LeftBorder = LineFormat
            oSheet.getCellRangeByName("A7:D7").BottomBorder = LineFormat
            oSheet.getCellRangeByName("A7:D7").VertJustify = 2
            oSheet.getCellRangeByName("A7:D7").HoriJustify = 2

            '---����� ����� ������
            i = 8
            Declarations.MyRec.MoveFirst()
            Declarations.MinMargin = Declarations.MyRec.Fields("MinMargin").Value
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("Str").Value
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("Code").Value
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("Name").Value
                oSheet.getCellRangeByName("D" & CStr(i)).Value = Declarations.MyRec.Fields("COST").Value
                LOFormatCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":D" & CStr(i), 4)

                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
            CheckNegativeMarginLO = False
            '----� ������ �����
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "$A$1"
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
            '----���������
            oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
            oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
        End If
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 10LO")
    End Function

    Public Function CheckNegativeProjectMargin(ByVal OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ��������� ������ 
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            Return CheckNegativeProjectMarginLO(OrderNum)
        Else
            Return CheckNegativeProjectMarginExcel(OrderNum)
        End If
    End Function

    Public Function CheckNegativeProjectMarginExcel(ByVal OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ��������� ������ 
        '// ����� � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer
        Dim MyScala As New SfwIII.Application

        On Error GoTo MyCatch

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)

        MySQLStr = "EXEC spp_ShippingCheck_GetNegativeProjectMarginInfo "
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"

        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CheckNegativeProjectMarginExcel = True
        Else
            '---���� ������������� �����, �� � �������
            MyObj = CreateObject("Excel.Application")
            MyObj.SheetsInNewWorkbook = 1
            MyWRKBook = MyObj.Workbooks.Add

            MyWRKBook.ActiveSheet.Columns("A:E").ColumnWidth = 15
            MyWRKBook.ActiveSheet.Columns("C").ColumnWidth = 40

            '---��������� �����
            MyWRKBook.ActiveSheet.Range("B1") = "����� �� ������� ����� " & OrderNum
            MyWRKBook.ActiveSheet.Range("A4") = "� ���� ������ ���������� ������ � ������ ���� �����������:"

            MyWRKBook.ActiveSheet.Range("B1").Select()
            MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
            MyWRKBook.ActiveSheet.Range("A4").Select()
            MyWRKBook.ActiveSheet.Range("A4").Font.Bold = True
            '---��������� �������� ������
            MyWRKBook.ActiveSheet.Range("A5") = "N ������"
            MyWRKBook.ActiveSheet.Range("B5") = "��� ��������"
            MyWRKBook.ActiveSheet.Range("C5") = "��� ��������"
            MyWRKBook.ActiveSheet.Range("D5") = "����"
            MyWRKBook.ActiveSheet.Range("E5") = "�������������"
            MyWRKBook.ActiveSheet.Range("F5") = "�����"
            MyWRKBook.ActiveSheet.Range("A5:F5").Select()
            MyWRKBook.ActiveSheet.Range("A5:F5").Borders(5).LineStyle = -4142
            MyWRKBook.ActiveSheet.Range("A5:F5").Borders(6).LineStyle = -4142
            With MyWRKBook.ActiveSheet.Range("A5:F5").Borders(7)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:F5").Borders(8)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:F5").Borders(9)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:F5").Borders(10)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:F5").Borders(11)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:F5").Interior
                .ColorIndex = 36
                .Pattern = 1
                .PatternColorIndex = -4105
            End With

            '---����� ����� ������
            i = 6
            Declarations.MyRec.MoveFirst()
            Declarations.MinMargin = Declarations.MyRec.Fields("MinMargin").Value
            While Not Declarations.MyRec.EOF
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)) = Declarations.MyRec.Fields("Str").Value
                MyWRKBook.ActiveSheet.Range("B" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" + CStr(i)) = Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("C" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" + CStr(i)) = Declarations.MyRec.Fields("Name").Value
                MyWRKBook.ActiveSheet.Range("D" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" + CStr(i)) = Declarations.MyRec.Fields("COST").Value
                MyWRKBook.ActiveSheet.Range("E" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("E" + CStr(i)) = Declarations.MyRec.Fields("PriCOST").Value
                MyWRKBook.ActiveSheet.Range("F" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("F" + CStr(i)) = IIf(Declarations.MyRec.Fields("COST").Value = 0, 0, IIf(Declarations.MyRec.Fields("PriCOST").Value = 0, 0, Math.Round((Declarations.MyRec.Fields("COST").Value - Declarations.MyRec.Fields("PriCOST").Value) / Declarations.MyRec.Fields("COST").Value * 100, 2)))
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While

            MyWRKBook.ActiveSheet.Range("A1").Select()
            MyObj.Application.Visible = True
            MyObj = Nothing
            CheckNegativeProjectMarginExcel = False
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 11")
    End Function

    Public Function CheckNegativeProjectMarginLO(ByVal OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ��� �� ������������� ����� � ��������� ������ 
        '// ����� � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim i As Integer

        On Error GoTo MyCatch

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)
        MySQLStr = "EXEC spp_ShippingCheck_GetNegativeProjectMarginInfo "
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CheckNegativeProjectMarginLO = True
        Else
            '---���� ������������� �����, �� � �������
            LOSetNotation(0)
            oServiceManager = CreateObject("com.sun.star.ServiceManager")
            oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
            oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
            Dim arg(1)
            arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
            arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
            oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
            oSheet = oWorkBook.getSheets().getByIndex(0)
            oFrame = oWorkBook.getCurrentController.getFrame
            '---������ �������
            oSheet.getColumns().getByName("A").Width = 2000
            oSheet.getColumns().getByName("B").Width = 5000
            oSheet.getColumns().getByName("C").Width = 10000
            oSheet.getColumns().getByName("D").Width = 4000
            oSheet.getColumns().getByName("E").Width = 4000
            oSheet.getColumns().getByName("F").Width = 4000
            '---��������� �����
            oSheet.getCellRangeByName("B1").String = "����� �� ������� ����� " & OrderNum
            oSheet.getCellRangeByName("A4").String = "� ���� ������ ���������� ������ � ������ ���� �����������:"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:B4", "Tahoma")
            LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1:B1")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1:B4", 11)
            '---��������� �������
            oSheet.getCellRangeByName("A5").String = "N ������"
            oSheet.getCellRangeByName("B5").String = "��� ��������"
            oSheet.getCellRangeByName("C5").String = "��� ��������"
            oSheet.getCellRangeByName("D5").String = "����"
            oSheet.getCellRangeByName("E5").String = "�������������"
            oSheet.getCellRangeByName("F5").String = "�����"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5:F5", "Calibri")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5:F5", 9)
            LOWrapText(oServiceManager, oDispatcher, oFrame, "A5:F5")
            oSheet.getCellRangeByName("A5:F5").CellBackColor = RGB(114, 251, 238)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
            Dim LineFormat As Object
            LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
            LineFormat.LineStyle = 0
            LineFormat.LineWidth = 70
            oSheet.getCellRangeByName("A5:F5").TopBorder = LineFormat
            oSheet.getCellRangeByName("A5:F5").RightBorder = LineFormat
            oSheet.getCellRangeByName("A5:F5").LeftBorder = LineFormat
            oSheet.getCellRangeByName("A5:F5").BottomBorder = LineFormat
            oSheet.getCellRangeByName("A5:F5").VertJustify = 2
            oSheet.getCellRangeByName("A5:F5").HoriJustify = 2

            '---����� ����� ������
            i = 6
            Declarations.MyRec.MoveFirst()
            Declarations.MinMargin = Declarations.MyRec.Fields("MinMargin").Value
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("Str").Value
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("Code").Value
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("Name").Value
                oSheet.getCellRangeByName("D" & CStr(i)).Value = Declarations.MyRec.Fields("COST").Value
                oSheet.getCellRangeByName("E" & CStr(i)).Value = Declarations.MyRec.Fields("PriCOST").Value
                oSheet.getCellRangeByName("F" & CStr(i)).Value = IIf(Declarations.MyRec.Fields("COST").Value = 0, 0, IIf(Declarations.MyRec.Fields("PriCOST").Value = 0, 0, Math.Round((Declarations.MyRec.Fields("COST").Value - Declarations.MyRec.Fields("PriCOST").Value) / Declarations.MyRec.Fields("COST").Value * 100, 2)))
                LOFormatCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":F" & CStr(i), 4)

                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
            CheckNegativeProjectMarginLO = False
            '----� ������ �����
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "$A$1"
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
            '----���������
            oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
            oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
        End If
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 11LO")
    End Function

    Public Function CheckRights(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - ����� �� ��������� ��� ������������ ������ � ����� ����������
        '// ������
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyPermission = False
            CheckRights = "���������"
        Else
            Declarations.MyPermission = True
            CheckRights = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 5")
        CheckRights = "���������"
    End Function

    Public Function CheckRights1(ByVal UserID As String, ByVal RoleName As String, ByVal MyOrder As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - ����� �� ��������� ��� ������ �����
        '// MarginType: 0 - ��������� �����, 1 - �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckRights1 = "���������"
            trycloseMyRec()
        Else
            trycloseMyRec()
            'MySQLStr = "SELECT MarginLevelTo "
            'MySQLStr = MySQLStr & "FROM tbl_MarginLimit WITH (NOLOCK) "
            'MySQLStr = MySQLStr & "WHERE (CheckLevel = N'" & RoleName & "')"
            '---��������� �� ������� ������
            MySQLStr = "SELECT tbl_MarginLimitMatrixDetails.MarginLevelTo "
            MySQLStr = MySQLStr & "FROM tbl_MarginLimitMatrixDetails INNER JOIN "
            MySQLStr = MySQLStr & "tbl_MarginLimitMatrix ON tbl_MarginLimitMatrixDetails.ID = tbl_MarginLimitMatrix.ID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON tbl_MarginLimitMatrix.ID = tbl_CustomerCard0300.MarginLimitLevel INNER JOIN "
            MySQLStr = MySQLStr & "OR010300 ON tbl_CustomerCard0300.SL01001 = OR010300.OR01003 "
            MySQLStr = MySQLStr & "WHERE (tbl_MarginLimitMatrixDetails.CheckLevel = N'" & RoleName & "') AND "
            MySQLStr = MySQLStr & "(OR010300.OR01001 = N'" & MyOrder & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                CheckRights1 = "���������"
                trycloseMyRec()
            Else
                Declarations.MyRec.MoveFirst()
                If Declarations.MyRec.Fields("MarginLevelTo").Value < Declarations.MinMargin Then
                    CheckRights1 = "���������"
                Else
                    CheckRights1 = "���������"
                End If
                trycloseMyRec()
            End If
        End If

        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 51")
        CheckRights1 = "���������"
    End Function

    Public Function SaveMarginInfo(ByVal MyGroup As String, ByVal MyOrder As String, ByVal MyUser As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ���������� �� �������� � ������ ���� �������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyUserID As Integer

        Declarations.MyMarginReason = ""
        While Declarations.MyMarginReason = ""
            MyLowMarginReason = New LowMarginReason
            MyLowMarginReason.ShowDialog()
        End While

        MySQLStr = "SELECT UserID "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (UPPER(UserName) = UPPER('" & MyUser & "')) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("������������ " & MyUser & " ��� � ��. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            MyUserID = Declarations.MyRec.Fields("UserID").Value
            trycloseMyRec()
            MySQLStr = "EXEC spp_ShippingCheck_SaveNegativeMarginInfo "
            MySQLStr = MySQLStr & "N'" & MyGroup & "', "
            MySQLStr = MySQLStr & "N'" & MyOrder & "', "
            MySQLStr = MySQLStr & CStr(MyUserID) & ", "
            MySQLStr = MySQLStr & "N'" & Declarations.MyMarginReason & "' "
            InitMyConn(True)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Function

    Public Function CheckShippingAllowed(ByVal OrderNumber As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ��� ���������� ������ ������ - ��������, ����� �� ������������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim CustomerType As String
        Dim MySQLStr As String

        On Error GoTo MyCatch
        Declarations.OrderID = OrderNumber
        '---��� ��� ��������� ����� ���� ���������� ����� ������
        InitMyConn(False)
        MySQLStr = "UPDATE OR01" & Declarations.CompanyID & "00 WITH(ROWLOCK) "
        MySQLStr = MySQLStr & "Set OR01024 = View_0.CorrSum "
        MySQLStr = MySQLStr & "FROM  OR01" & Declarations.CompanyID & "00 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT OR01" & Declarations.CompanyID & "00_1.OR01001, "
        '��� ����������
        'MySQLStr = MySQLStr & "SUM(ROUND(OR03" & Declarations.CompanyID & "00.OR03008 - OR03" & Declarations.CompanyID & "00.OR03008 * OR03" & Declarations.CompanyID & "00.OR03018 / 100, 2) "
        'MySQLStr = MySQLStr & "* OR03" & Declarations.CompanyID & "00.OR03011) AS CorrSum "
        '� ����������� ���� � ������ �� 2 �����
        'MySQLStr = MySQLStr & "SUM(Round(Round((OR03" & Declarations.CompanyID & "00.OR03008 - "
        'MySQLStr = MySQLStr & "OR03" & Declarations.CompanyID & "00.OR03008 * OR03" & Declarations.CompanyID & "00.OR03018 / 100), 2) "
        'MySQLStr = MySQLStr & "* OR03" & Declarations.CompanyID & "00.OR03011, 2)) As CorrSum "
        MySQLStr = MySQLStr & "SUM(ROUND(ROUND(OR030300.OR03008 * "
        MySQLStr = MySQLStr & "(100 - CONVERT(float,OR030300.OR03018) - "
        MySQLStr = MySQLStr & "CONVERT(float,OR030300.OR03017)) / 100, 2) * OR030300.OR03011 / OR030300.OR03022, 2)) As CorrSum "

        MySQLStr = MySQLStr & "FROM OR01" & Declarations.CompanyID & "00 AS OR010300_1 INNER JOIN "
        MySQLStr = MySQLStr & "OR03" & Declarations.CompanyID & "00 ON OR01" & Declarations.CompanyID & "00_1.OR01001 = OR03" & Declarations.CompanyID & "00.OR03001 "
        MySQLStr = MySQLStr & "WHERE (OR03" & Declarations.CompanyID & "00.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY OR01" & Declarations.CompanyID & "00_1.OR01001 "
        MySQLStr = MySQLStr & "ORDER BY OR01" & Declarations.CompanyID & "00_1.OR01001) AS View_0 ON OR01" & Declarations.CompanyID & "00.OR01001 = View_0.OR01001 "
        MySQLStr = MySQLStr & "WHERE (OR01" & Declarations.CompanyID & "00.OR01001 = N'" & OrderNumber & "') "
        Declarations.MyConn.Execute(MySQLStr)

        '---����������� ���������� ����� ����������������� (���������
        MySQLStr = "Update OR030300 "
        MySQLStr = MySQLStr & "SET OR03014 = CONVERT(int, SL010300.SL01032) + CONVERT(int, SC010300.SC01035) "
        MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR030300.OR03119 = SL010300.SL01001 INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03001 = N'" & OrderNumber & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---����������� ���������� ����� ���
        '---� ��������� ������
        'MySQLStr = "Update OR010300 "
        'MySQLStr = MySQLStr & "Set OR01095 = Right('00' + Ltrim(Rtrim(SL010300.SL01107)),2), "
        'MySQLStr = MySQLStr & "OR01118 = Right('00' + Ltrim(Rtrim(SL010300.SL01107)),2) "
        'MySQLStr = MySQLStr & "FROM OR010300 INNER JOIN "
        'MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        'MySQLStr = MySQLStr & "WHERE (OR010300.OR01001 = N'" & OrderNumber & "') "
        MySQLStr = "Update OR010300 "
        MySQLStr = MySQLStr & "Set OR01093 = N'', "
        MySQLStr = MySQLStr & "OR01094 = N'', "
        MySQLStr = MySQLStr & "OR01095 = N'', "
        MySQLStr = MySQLStr & "OR01118 = N'' "
        MySQLStr = MySQLStr & "FROM OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01001 = N'" & OrderNumber & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---� ������� ������
        'MySQLStr = "Update OR030300 "
        'MySQLStr = MySQLStr & "Set OR03061 = CONVERT(nvarchar, CONVERT(int, SL010300.SL01107)) "
        'MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
        'MySQLStr = MySQLStr & "SL010300 ON OR030300.OR03119 = SL010300.SL01001 "
        'MySQLStr = MySQLStr & "WHERE (OR030300.OR03001 = N'" & OrderNumber & "') "
        MySQLStr = "Update OR030300 "
        MySQLStr = MySQLStr & "Set OR03061 = SC010300.SC01144 "
        MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03001 = N'" & OrderNumber & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---����������� ���������� ������� ������ � ������� � ������������ � ����������
        MySQLStr = "UPDATE OR030300 "
        MySQLStr = MySQLStr & "Set OR03028 = OR010300.OR01025 "
        MySQLStr = MySQLStr & "FROM OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "OR030300 ON OR010300.OR01001 = OR030300.OR03001 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01001 = N'" & OrderNumber & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)


        CustomerType = CheckCustomerType(OrderNumber)
        '-------��������� ����������---------------------------------------------------------
        If CustomerType = "���������" Then
            If CheckCreditLimit(OrderNumber, Declarations.CustomerID) = False Then
                CheckShippingAllowed = False
                Exit Function
            Else
                CheckShippingAllowed = True
                Exit Function
            End If
            '-------�����������------------------------------------------------------------------
        ElseIf CustomerType = "�����������" Then
            If CheckOrderPayment(OrderNumber) = False Then
                CheckShippingAllowed = False
                Exit Function
            Else
                CheckShippingAllowed = True
                Exit Function
            End If
            '-------���������� �����������-------------------------------------------------------
        ElseIf CustomerType = "����������" Then
            CheckShippingAllowed = True
            Exit Function
            '-------��������� ����������---------------------------------------------------------
        ElseIf CustomerType = "���������" Then
            If CheckOrderPayment(OrderNumber) = False Then
                CheckShippingAllowed = False
                Exit Function
            Else
                CheckShippingAllowed = True
                Exit Function
            End If
            '-------����������� �� - �� ����-----------------------------------------------------
        ElseIf CustomerType = "���������������������" Then
            MsgBox("� ���������� ���� ���������� ��������� ��� �������. ��������� ��������, ���� ������ ���������� - ���������� � �������������� �������.", vbCritical, "��������!")
            CheckShippingAllowed = False
            Exit Function
            '-------������ ����������� � ������ ��������-----------------------------------------
        Else
            MsgBox("������, ��������� � ������ ������, ����������� � ������� ��������. ���������� � �������������� �������.", vbCritical, "��������!")
            CheckShippingAllowed = False
            Exit Function
        End If
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 3")
        CheckShippingAllowed = False
    End Function

    Public Function CheckCreditLimit(ByVal OrderNumber As String, ByVal CustomerID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���������� ������ ������������ - ����� �� ��������� �����
        '// �������� �������������� ����� �������� � ����� (� �������� ����� ������� ����� � ������)
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim PrepaymentSum As Double         '---����� ����� ���������
        Dim CardPayedSum As Double          '---���������� �� �������� � WEB �����

        Declarations.IsWEBOrder = CheckWEBOrNot(OrderNumber)
        CardPayedSum = GetCardPayment(OrderNumber)

        On Error GoTo MyCatch
        'MySQLStr = "SELECT OR01002 AS OrderType, OR01028, "
        'MySQLStr = MySQLStr & "CASE OR01028 WHEN 0 THEN OR01024 * 1.18 ELSE OR01024 * OR01067 * 1.18 END AS OrderSum "
        'MySQLStr = MySQLStr & "FROM dbo.OR010300 "
        'MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Right("0000000000" & OrderNumber, 10) & "')"
        MySQLStr = "SELECT OR010300_1.OR01002 AS OrderType, OR010300_1.OR01028, View_1.OrderSum "
        MySQLStr = MySQLStr & "FROM (SELECT OR030300.OR03001, "
        MySQLStr = MySQLStr & "ROUND(SUM(ROUND(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 "
        MySQLStr = MySQLStr & "END) * (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 / OR030300.OR03022, 2)) "
        MySQLStr = MySQLStr & "+ SUM(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 END) "
        MySQLStr = MySQLStr & "* (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 * CONVERT(float, SY290300.SY29003) / 100 / OR030300.OR03022), 2) "
        MySQLStr = MySQLStr & "AS OrderSum "
        MySQLStr = MySQLStr & "FROM  OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 ON OR030300.OR03001 = OR010300.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SY290300 ON OR030300.OR03061 = SY290300.SY29001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY OR030300.OR03001 "
        MySQLStr = MySQLStr & "HAVING (OR030300.OR03001 = N'" & Right("0000000000" & OrderNumber, 10) & "')) AS View_1 INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 AS OR010300_1 ON View_1.OR03001 = OR010300_1.OR01001 "
        MySQLStr = MySQLStr & "GROUP BY OR010300_1.OR01002, OR010300_1.OR01028, View_1.OrderSum "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then '---��� �������� - �� ���������
            trycloseMyRec()
            CheckCreditLimit = True
            Exit Function
        End If
        If Declarations.MyRec.Fields("OrderType").Value = "8" Or Declarations.MyRec.Fields("OrderType").Value = "2" Then '---������� �� ��������� � 2 ���
            trycloseMyRec()
            CheckCreditLimit = True
            Exit Function
        ElseIf Declarations.MyRec.Fields("OrderType").Value = "0" Or Declarations.MyRec.Fields("OrderType").Value = "1" _
            Or Declarations.MyRec.Fields("OrderType").Value = "4" Or Declarations.MyRec.Fields("OrderType").Value = "5" Then '---0, 1, 4 ��� 5 ��� ������
            Declarations.OrderSum = Declarations.MyRec.Fields("OrderSum").Value  '---����� � ���
            Declarations.CurrCode = Declarations.MyRec.Fields("OR01028").Value
            'Else '--- 4 (���� ��������) ��� 5
            '    trycloseMyRec
            '    MySQLStr = "SELECT OR20002 AS OrderType, "
            '    MySQLStr = MySQLStr & "CASE OR20028 WHEN 0 THEN OR20024 * 1.18 ELSE OR20024 * OR20067 * 1.18 END AS OrderSum, "
            '    MySQLStr = MySQLStr & "OR20028 "
            '    MySQLStr = MySQLStr & "FROM dbo.OR200300 "
            '    MySQLStr = MySQLStr & "WHERE (OR20002 = N'1') "
            '    MySQLStr = MySQLStr & "AND (OR20001 = N'" & Right("0000000000" & OrderNumber, 10) & "')"
            '    InitMyConn True
            '    InitMyRec True, MySQLStr
            '    If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
            '        MsgBox "� ������� �� ������ ����� 1 ����, ��������������� ���������� ������ 4 ����. ���������� � �������������� ������� Scala.", vbCritical, "��������!"
            '        trycloseMyRec
            '        CheckCreditLimit = False
            '        Exit Function
            '    End If
            '    declarations.OrderSum = declarations.MyRec!OrderSum   '---� ��� � ������
            '    declarations.CurrCode = declarations.MyRec!OR20028
        End If
        trycloseMyRec()

        '---������ 1 ����
        MySQLStr = "SELECT CustomerCode, SUM(Sum) AS Avance1Type "
        MySQLStr = MySQLStr & "FROM (SELECT TOP (100) PERCENT MIN(dbo.SL030300.SL03001) AS CustomerCode, "
        MySQLStr = MySQLStr & "SUM(dbo.SL210300.SL21007) AS Sum "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.OR010300 ON RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = OR010300.OR01001 INNER JOIN "
        MySQLStr = MySQLStr & "dbo.SL210300 ON dbo.SL030300.SL03001 = dbo.SL210300.SL21001 AND "
        MySQLStr = MySQLStr & "dbo.SL030300.SL03002 = dbo.SL210300.SL21002 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR010300.OR01002 IN ('0', '1')) AND "
        MySQLStr = MySQLStr & "(ABS(dbo.SL030300.SL03013) < 0.01) "
        MySQLStr = MySQLStr & "GROUP BY dbo.SL030300.SL03002, dbo.OR010300.OR01028 "
        MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21007)) > 0.01) AND "
        MySQLStr = MySQLStr & "(MIN(dbo.SL030300.SL03001) = N'" & CustomerID & "')) AS t1 "
        MySQLStr = MySQLStr & "GROUP BY CustomerCode"
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.Avance1Type = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.Avance1Type = Declarations.MyRec.Fields("Avance1Type").Value
        End If
        trycloseMyRec()

        '---������ 2 ����
        MySQLStr = "SELECT CustomerCode, SUM(Sum) AS Avance2Type "
        MySQLStr = MySQLStr & "FROM (SELECT TOP (100) PERCENT MIN(dbo.SL030300.SL03001) AS CustomerCode, "
        MySQLStr = MySQLStr & "SUM(dbo.SL210300.SL21007) AS Sum "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.OR200300 ON RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = OR200300.OR20001 INNER JOIN "
        MySQLStr = MySQLStr & "dbo.SL210300 ON dbo.SL030300.SL03001 = dbo.SL210300.SL21001 AND "
        MySQLStr = MySQLStr & "dbo.SL030300.SL03002 = dbo.SL210300.SL21002 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR200300.OR20002 IN ('0', '1')) AND "
        MySQLStr = MySQLStr & "(ABS(dbo.SL030300.SL03013) < 0.01) "
        MySQLStr = MySQLStr & "GROUP BY dbo.SL030300.SL03002 "
        MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21007)) > 0.01) AND "
        MySQLStr = MySQLStr & "(MIN(dbo.SL030300.SL03001) = N'" & CustomerID & "')) AS t1 "
        MySQLStr = MySQLStr & "GROUP BY CustomerCode"
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.Avance2Type = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.Avance2Type = Declarations.MyRec.Fields("Avance2Type").Value
        End If
        trycloseMyRec()

        '---�������� ����� ����������
        If CardPayedSum > (Declarations.Avance1Type + Declarations.Avance2Type) Then
            PrepaymentSum = CardPayedSum
        Else
            PrepaymentSum = Declarations.Avance1Type + Declarations.Avance2Type
        End If

        '---����� ������������� ������
        '---�� ������ - ��������
        MySQLStr = "SELECT t11.CustomerCode, ISNULL(t11.InvoiceSum,0) as InvoiceSum, ISNULL(t12.PayedSum,0) as PayedSum "
        MySQLStr = MySQLStr & "FROM (SELECT SL03001 AS CustomerCode, SUM(SL03013) AS InvoiceSum "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(LEFT(SL03017, 6) = '621010') "
        MySQLStr = MySQLStr & "GROUP BY SL03001) AS t11 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SL21001 AS CustomerCode, SUM(SL21007) AS PayedSum "
        MySQLStr = MySQLStr & "FROM dbo.SL210300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL21006 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(SL21002 IN "
        MySQLStr = MySQLStr & "(SELECT DISTINCT SL03002 "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 AS SL030300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(LEFT(SL03017, 6) = '621010'))) "
        MySQLStr = MySQLStr & "GROUP BY SL21001) AS t12 ON t11.CustomerCode = t12.CustomerCode "
        MySQLStr = MySQLStr & "WHERE (t11.CustomerCode = N'" & CustomerID & "') "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.InvoiceDebt = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.InvoiceDebt = Declarations.MyRec.Fields("InvoiceSum").Value - Declarations.MyRec.Fields("PayedSum").Value
        End If
        trycloseMyRec()

        '---����������� � �������� ������ (����� ����� ������ - ������)
        'MySQLStr = "SELECT ISNULL(SUM(OrderSum),0) AS OrderSum "
        'MySQLStr = MySQLStr & "FROM (SELECT OR01003 AS CustomerCode, "
        'MySQLStr = MySQLStr & "SUM(CASE OR01028 WHEN '0' THEN OR01024 * 1.18 WHEN '00' THEN OR01024 * 1.18 ELSE OR01024 * OR01067 * 1.18 END) "
        'MySQLStr = MySQLStr & "AS OrderSum "
        'MySQLStr = MySQLStr & "FROM dbo.OR010300 "
        'MySQLStr = MySQLStr & "WHERE ((OR01002 = 1) AND (OR01008 = 3)) OR "
        'MySQLStr = MySQLStr & "((OR01002 = 4) AND (OR01008 = 3)) "
        'MySQLStr = MySQLStr & "GROUP BY OR01003, OR01028 "
        'MySQLStr = MySQLStr & "HAVING (OR01003 = N'" & CustomerID & "')) AS t1 "
        MySQLStr = "SELECT ISNULL(SUM(View_1.OrderSum), 0) AS OrderSum "
        MySQLStr = MySQLStr & "FROM (SELECT OR030300.OR03001, "
        MySQLStr = MySQLStr & "ROUND(SUM(ROUND(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 "
        MySQLStr = MySQLStr & "END) * (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 / OR030300.OR03022, 2)) "
        MySQLStr = MySQLStr & "+ SUM(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 END) "
        MySQLStr = MySQLStr & "* (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 * CONVERT(float, SY290300.SY29003) / 100 / OR030300.OR03022), 2) "
        MySQLStr = MySQLStr & "AS OrderSum "
        MySQLStr = MySQLStr & "FROM OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 ON OR030300.OR03001 = OR010300.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SY290300 ON OR030300.OR03061 = SY290300.SY29001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY OR030300.OR03001) AS View_1 INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 AS OR010300_1 ON View_1.OR03001 = OR010300_1.OR01001 "
        MySQLStr = MySQLStr & "WHERE (OR010300_1.OR01001 <> N'" & Right("0000000000" & OrderNumber, 10) & "') "
        MySQLStr = MySQLStr & "AND (OR010300_1.OR01003 = N'" & CustomerID & "') "
        MySQLStr = MySQLStr & "AND (OR010300_1.OR01002 = 1 OR "
        MySQLStr = MySQLStr & "OR010300_1.OR01002 = 4) AND (OR010300_1.OR01008 = 3) "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.OrderDebt = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.OrderDebt = Declarations.MyRec.Fields("OrderSum").Value
        End If
        trycloseMyRec()

        '--��������� - �� ���������� �� ������ ������������ � ������
        MySQLStr = "SELECT COUNT(InvoiceNum) AS OverduePaymentQTY, "
        MySQLStr = MySQLStr & "ISNULL(SUM(Overdue),0) AS Overdue "
        MySQLStr = MySQLStr & "FROM (SELECT t11.InvoiceNum, "
        MySQLStr = MySQLStr & "t11.InvoiceSum - ISNULL(t12.PayedSum, 0) AS Overdue "
        MySQLStr = MySQLStr & "FROM (SELECT TOP 100 PERCENT SL03001 AS CustomerCode, "
        MySQLStr = MySQLStr & "SL03002 AS InvoiceNum, SUM(SL03013) AS InvoiceSum, "
        MySQLStr = MySQLStr & "SL03006 As InvoiceData "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(LEFT(SL03017, 6) = '621010') "
        MySQLStr = MySQLStr & "GROUP BY SL03001, SL03002, SL03006 "
        MySQLStr = MySQLStr & "HAVING (SL03001 = N'" & CustomerID & "') "
        MySQLStr = MySQLStr & "ORDER BY InvoiceNum) AS t11 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP 100 PERCENT SL21001 AS CustomerCode, "
        MySQLStr = MySQLStr & "SL21002 AS InvoiceNum, "
        MySQLStr = MySQLStr & "SUM(SL21007) AS PayedSum "
        MySQLStr = MySQLStr & "FROM dbo.SL210300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL21006 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(SL21002 IN (SELECT DISTINCT SL03002 "
        MySQLStr = MySQLStr & "FROM dbo.SL030300 AS SL030300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03005 <= GETDATE()) AND "
        MySQLStr = MySQLStr & "(LEFT(SL03017, 6) = '621010'))) "
        MySQLStr = MySQLStr & "GROUP BY SL21001, SL21002 "
        MySQLStr = MySQLStr & "HAVING (SL21001 = N'" & CustomerID & "') "
        MySQLStr = MySQLStr & "ORDER BY InvoiceNum) AS t12 ON t11.InvoiceNum = t12.InvoiceNum AND "
        MySQLStr = MySQLStr & "t11.CustomerCode = t12.CustomerCode "
        MySQLStr = MySQLStr & "WHERE (t11.InvoiceData < dateadd(d, -1, GETDATE())) AND "
        MySQLStr = MySQLStr & "(ISNULL(t12.PayedSum, 0) < t11.InvoiceSum - 1) "
        MySQLStr = MySQLStr & "GROUP BY DATEDIFF(day, t11.InvoiceData, GETDATE()), "
        MySQLStr = MySQLStr & "t11.InvoiceNum, t11.InvoiceSum, t12.PayedSum) AS t13 "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        Declarations.OverduePaymentQTY = Declarations.MyRec.Fields("OverduePaymentQTY").Value
        Declarations.Overdue = Declarations.MyRec.Fields("Overdue").Value
        trycloseMyRec()

        '---���������� � ������
        If (Declarations.CreditAmount = 0 And Declarations.OverduePaymentQTY = 0) Or _
            ((Declarations.CreditAmount + PrepaymentSum - _
            Declarations.OrderSum - Declarations.InvoiceDebt - Declarations.OrderDebt) > -1 And _
            Declarations.OverduePaymentQTY = 0) Then
            '---���������� ����� ������� � ������ ������ ���
            '---���� � ������ ����� ����� 80% ���������� ������ - �������������
            If Declarations.CreditAmount <> 0 Then
                If (Declarations.OrderSum + Declarations.InvoiceDebt + Declarations.OrderDebt - PrepaymentSum) / Declarations.CreditAmount > 0.8 Then
                    MsgBox("������ ������ ����� 80% ���������� ������")
                End If
            End If
            CheckCreditLimit = True
            Exit Function
        Else
            '---���� ���������� ����� ������� � (���) ������ ������ �������
            CheckRights(Declarations.UserCode, "ShipmentsOverLimit")
            Declarations.CmdToShip = False
            '---�������� ���� ��������� ���������� � ������ ������� �� �������� ����� ������
            MyCreditDialog = New CreditDialog
            MyCreditDialog.ShowDialog()
            If Declarations.CmdToShip = True And Declarations.MyPermission = True Then
                SaveDataAboutOverCreditShipments(Declarations.UserCode, Declarations.OrderID, _
                    Declarations.OrderSum, Declarations.Avance1Type, Declarations.Avance2Type, _
                    Declarations.InvoiceDebt, Declarations.OrderDebt, Declarations.Overdue)
                CheckCreditLimit = True
                Exit Function
            End If
        End If
        CheckCreditLimit = False
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 6")
        CheckCreditLimit = False
    End Function

    Public Function SaveDataAboutOverCreditShipments(ByVal UserName As String, ByVal OrderID As String, ByVal OrderSum As Double, _
    ByVal Avance1Type As Double, ByVal Avance2Type As Double, ByVal InvoiceDebt As Double, ByVal OrderDebt As Double, ByVal Overdue As Double)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ����� � �������� ���������� - ��� �������� �������� ����� ���������� ������
        '// � ��������� ������ � �������� �� ��� ������
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "INSERT INTO tbl_CreditShipmentsOverLimit"
        MySQLStr = MySQLStr & "(UserName, "
        MySQLStr = MySQLStr & "OrderID, "
        MySQLStr = MySQLStr & "ActDate, "
        MySQLStr = MySQLStr & "Curr, "
        MySQLStr = MySQLStr & "OrderSum, "
        MySQLStr = MySQLStr & "Avance1Type, "
        MySQLStr = MySQLStr & "Avance2Type, "
        MySQLStr = MySQLStr & "InvoiceDebt, "
        MySQLStr = MySQLStr & "OrderDebt, "
        MySQLStr = MySQLStr & "CreditAmount,"
        MySQLStr = MySQLStr & "CreditInDays,"
        MySQLStr = MySQLStr & "Comments, "
        MySQLStr = MySQLStr & "Overdue) "
        MySQLStr = MySQLStr & "VALUES (N'" & UserName & "', "
        MySQLStr = MySQLStr & "N'" & OrderID & "', "
        MySQLStr = MySQLStr & "GETDATE ( ), "
        MySQLStr = MySQLStr & "0, "
        MySQLStr = MySQLStr & Replace(CStr(OrderSum), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Avance1Type), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Avance2Type), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(InvoiceDebt), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(OrderDebt), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Declarations.CreditAmount), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Declarations.CreditInDays), ",", ".") & ", N'"
        MySQLStr = MySQLStr & Declarations.MyReason & "', "
        MySQLStr = MySQLStr & Replace(CStr(Overdue), ",", ".") & ") "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 9")
    End Function

    Public Function CheckOrderPayment(ByVal OrderNumber As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� - ��� �� ������ �� ����� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim Reply As Object
        Dim PrepaymentSum As Double         '---����� ����� ���������
        Dim CardPayedSum As Double          '---���������� �� �������� � WEB �����

        Declarations.IsWEBOrder = CheckWEBOrNot(OrderNumber)
        CardPayedSum = GetCardPayment(OrderNumber)

        On Error GoTo MyCatch
        '--- ��� ������� ����� ������--------------------------------------------------------
        'MySQLStr = "SELECT OR01002, OR01024, OR01028, OR01036 "
        'MySQLStr = MySQLStr & "FROM OR010300 "
        'MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Right("0000000000" & OrderNumber, 10) & "')"
        MySQLStr = "SELECT OR010300_1.OR01002 AS OrderType, OR010300_1.OR01028, View_1.OrderSum "
        MySQLStr = MySQLStr & "FROM (SELECT OR030300.OR03001, "
        MySQLStr = MySQLStr & "ROUND(SUM(ROUND(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 "
        MySQLStr = MySQLStr & "END) * (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 / OR030300.OR03022, 2)) "
        MySQLStr = MySQLStr & "+ SUM(ROUND((OR030300.OR03008 * "
        MySQLStr = MySQLStr & "CASE WHEN OR010300.OR01067 = 0 THEN 1 ELSE OR010300.OR01067 END) "
        MySQLStr = MySQLStr & "* (100 - CONVERT(float,OR030300.OR03018) - CONVERT(float,OR030300.OR03017)) / 100, 2) * "
        MySQLStr = MySQLStr & "OR030300.OR03011 * CONVERT(float, SY290300.SY29003) / 100 / OR030300.OR03022), 2) "
        MySQLStr = MySQLStr & "AS OrderSum "
        MySQLStr = MySQLStr & "FROM  OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 ON OR030300.OR03001 = OR010300.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SY290300 ON OR030300.OR03061 = SY290300.SY29001 "
        MySQLStr = MySQLStr & "WHERE (OR030300.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY OR030300.OR03001 "
        MySQLStr = MySQLStr & "HAVING (OR030300.OR03001 = N'" & Right("0000000000" & OrderNumber, 10) & "')) AS View_1 INNER JOIN "
        MySQLStr = MySQLStr & "OR010300 AS OR010300_1 ON View_1.OR03001 = OR010300_1.OR01001 "
        MySQLStr = MySQLStr & "GROUP BY OR010300_1.OR01002, OR010300_1.OR01028, View_1.OrderSum "

        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then '---��� �������� - �� ���������
            trycloseMyRec()
            CheckOrderPayment = True
            Exit Function
        End If
        If Declarations.MyRec.Fields("OrderType").Value = "8" Or Declarations.MyRec.Fields("OrderType").Value = "2" Then '---������� �� ��������� � 2 ���
            trycloseMyRec()
            CheckOrderPayment = True
            Exit Function
        ElseIf Declarations.MyRec.Fields("OrderType").Value = "0" Or Declarations.MyRec.Fields("OrderType").Value = "1" Or Declarations.MyRec.Fields("OrderType").Value = "4" Then '---0, 4 ��� 1 ��� ������
            Declarations.OrderSum = Declarations.MyRec.Fields("OrderSum").Value  '---� ���
            Declarations.CurrCode = 0 'Declarations.MyRec!OR01028
        Else '--- 4 (���� ��������) ��� 5
            trycloseMyRec()
            'MySQLStr = "SELECT OR20002, OR20024, OR20028, OR20036 "
            'MySQLStr = MySQLStr & "FROM OR200300 "
            'MySQLStr = MySQLStr & "WHERE (OR20001 = N'" & Right("0000000000" & OrderNumber, 10) & "') "
            'MySQLStr = MySQLStr & "AND (OR20002 = N'1')"
            'InitMyConn True
            'InitMyRec True, MySQLStr
            'If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '    MsgBox "� ������� �� ������ ����� 1 ����, ��������������� ���������� ������ 4 ����. ���������� � �������������� ������� Scala.", vbCritical, "��������!"
            '    trycloseMyRec
            '    CheckOrderPayment = False
            '    Exit Function
            'End If
            'Declarations.OrderSum = Declarations.MyRec!OR20024 * 1.18   '---� ���
            'Declarations.CurrCode = Declarations.MyRec!OR20028
        End If
        trycloseMyRec()

        '---������ 1 ����
        MySQLStr = "SELECT TOP 100 PERCENT RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) AS OrderNumber, "
        If Declarations.CurrCode = "0" Or Declarations.CurrCode = "00" Then '---�����
            MySQLStr = MySQLStr & "ISNULL(SUM(dbo.SL210300.SL21007),0) As SUM "
        Else                                      '--- ������
            MySQLStr = MySQLStr & "ISNULL(SUM(dbo.SL210300.SL21008),0) As SUM "
        End If
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.OR010300 ON RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = dbo.OR010300.OR01001 "
        MySQLStr = MySQLStr & "INNER JOIN dbo.SL210300 ON dbo.SL030300.SL03001 = dbo.SL210300.SL21001 AND "
        MySQLStr = MySQLStr & "dbo.SL030300.SL03002 = dbo.SL210300.SL21002 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR010300.OR01002 IN ('0', '1')) AND (ABS(dbo.SL030300.SL03013) < 0.01) "
        MySQLStr = MySQLStr & "GROUP BY dbo.SL030300.SL03002 "
        If CurrCode = "0" Or CurrCode = "00" Then '---�����
            MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21007)) > 0.01) AND "
        Else                                      '--- ������
            MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21008)) > 0.01) AND "
        End If
        MySQLStr = MySQLStr & "(RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = '" & Right("0000000000" & OrderNumber, 10) & "')"
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.Avance1Type = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.Avance1Type = Declarations.MyRec.Fields("SUM").Value
        End If
        trycloseMyRec()
        '---������ 0 ����
        MySQLStr = "SELECT  TOP 100 PERCENT RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) AS OrderNumber, "
        If Declarations.CurrCode = "0" Or Declarations.CurrCode = "00" Then '---�����
            MySQLStr = MySQLStr & "ISNULL(SUM(dbo.SL210300.SL21007),0) AS Sum "
        Else                                      '--- ������
            MySQLStr = MySQLStr & "ISNULL(SUM(dbo.SL210300.SL21008),0) AS Sum "
        End If
        MySQLStr = MySQLStr & "FROM dbo.SL030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "dbo.OR200300 ON RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = dbo.OR200300.OR20001 "
        MySQLStr = MySQLStr & "INNER JOIN dbo.SL210300 ON dbo.SL030300.SL03001 = dbo.SL210300.SL21001 AND "
        MySQLStr = MySQLStr & "dbo.SL030300.SL03002 = dbo.SL210300.SL21002 "
        MySQLStr = MySQLStr & "WHERE (dbo.OR200300.OR20002 IN ('0', '1')) AND (ABS(dbo.SL030300.SL03013) < 0.01) "
        MySQLStr = MySQLStr & "GROUP BY dbo.SL030300.SL03002 "
        If CurrCode = "0" Or CurrCode = "00" Then '---�����
            MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21007)) > 0.01) AND "
        Else                                      '--- ������
            MySQLStr = MySQLStr & "HAVING (ABS(SUM(dbo.SL210300.SL21008)) > 0.01) AND "
        End If
        MySQLStr = MySQLStr & "(RIGHT('0000000000' + LEFT(LTRIM(RTRIM(SL030300.SL03002)), dbo.fnc_GetFirstNotDigitPos(LTRIM(RTRIM(SL030300.SL03002))) - 1),10) = '" & Right("0000000000" & OrderNumber, 10) & "')"
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.Avance2Type = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.Avance2Type = Declarations.MyRec.Fields("SUM").Value
        End If
        trycloseMyRec()

        '---�������� ����� ����������
        If CardPayedSum > (Declarations.Avance1Type + Declarations.Avance2Type) Then
            PrepaymentSum = CardPayedSum
        Else
            PrepaymentSum = Declarations.Avance1Type + Declarations.Avance2Type
        End If

        '---������
        'MySQLStr = "SELECT SL03036 AS OrderNumber, "
        'If Declarations.CurrCode = "0" Or Declarations.CurrCode = "00" Then '---�����
        '    MySQLStr = MySQLStr & "ISNULL(SUM(SL03013),0) AS Sum "
        'Else                                      '--- ������
        '    MySQLStr = MySQLStr & "ISNULL(SUM(SL03014),0) AS Sum "
        'End If
        'MySQLStr = MySQLStr & "FROM dbo.SL030300 "
        'MySQLStr = MySQLStr & "GROUP BY SL03036 "
        'MySQLStr = MySQLStr & "HAVING (SL03036 = N'" & Right("0000000000" & OrderNumber, 10) & "')"
        'InitMyConn True
        'InitMyRec True, MySQLStr
        'If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Declarations.MyPayment = 0
        'Else
        '    Declarations.MyRec.MoveFirst
        '    Declarations.MyPayment = Declarations.MyRec!SUM
        'End If
        'trycloseMyRec

        '---���� ���� ������������� �� ������ ���������
        MySQLStr = "SELECT CASE WHEN View_1.InvoiceSum - View_1.PayedSum < 10 THEN 0 ELSE View_1.InvoiceSum - View_1.PayedSum END AS Debt "
        MySQLStr = MySQLStr & "FROM OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT t11.CustomerCode, ISNULL(t11.InvoiceSum, 0) AS InvoiceSum, "
        MySQLStr = MySQLStr & "ISNULL(t12.PayedSum, 0) AS PayedSum "
        MySQLStr = MySQLStr & "FROM (SELECT SL03001 AS CustomerCode, SUM(SL03013) AS InvoiceSum "
        MySQLStr = MySQLStr & "FROM SL030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03004 <= GETDATE()) AND (LEFT(SL03017, 6) = '621010') "
        MySQLStr = MySQLStr & "GROUP BY SL03001) AS t11 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SL21001 AS CustomerCode, SUM(SL21007) AS PayedSum "
        MySQLStr = MySQLStr & "FROM SL210300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL21006 <= GETDATE()) AND (SL21002 IN "
        MySQLStr = MySQLStr & "(SELECT DISTINCT SL03002 "
        MySQLStr = MySQLStr & "FROM SL030300 AS SL030300_1 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL03004 <= GETDATE()) AND (LEFT(SL03017, 6) = '621010'))) "
        MySQLStr = MySQLStr & "GROUP BY SL21001) AS t12 ON t11.CustomerCode = t12.CustomerCode) AS View_1 "
        MySQLStr = MySQLStr & "ON OR010300.OR01003 = View_1.CustomerCode "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01001 = N'" & Right("0000000000" & OrderNumber, 10) & "') "
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.InvoiceDebt = 0
        Else
            Declarations.MyRec.MoveFirst()
            Declarations.InvoiceDebt = Declarations.MyRec.Fields("Debt").Value
        End If
        trycloseMyRec()


        '---���������� � ������
        If (Declarations.OrderSum - Declarations.MyPayment - PrepaymentSum) < 1 And Declarations.InvoiceDebt = 0 Then
            '---����� �������
            CheckOrderPayment = True
            Exit Function
        Else
            '---����� �� ������� ��� ������� �� ���������
            CheckRights(Declarations.UserCode, "ShipmentsOverLimit")
            Declarations.CmdToShip = False
            '---�������� ���� ��������� ���������� � ������ ������� �� �������� ����� ������
            MyNonCreditDialog = New NonCreditDialog
            MyNonCreditDialog.ShowDialog()
            If Declarations.CmdToShip = True And Declarations.MyPermission = True Then
                SaveDataAboutNotPayedShipments(Declarations.UserCode, Declarations.OrderID, _
                    Declarations.OrderSum, Declarations.MyPayment, Declarations.Avance1Type, Declarations.Avance2Type)
                CheckOrderPayment = True
                Exit Function
            End If
        End If
        CheckOrderPayment = False
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 7")
        CheckOrderPayment = False
    End Function

    Public Function SaveDataAboutNotPayedShipments(ByVal UserName As String, ByVal OrderID As String, ByVal OrderSum As Double, _
    ByVal MyPayment As Double, ByVal Avance1Type As Double, ByVal Avance2Type As Double)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ����� � �������� ���������� - ��� �������� ������������ ��������
        '// � ��������� ������ � �������� �� ��� ������
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "INSERT INTO tbl_NonCreditShipmentsOverLimit"
        MySQLStr = MySQLStr & "(UserName, "
        MySQLStr = MySQLStr & "OrderID, "
        MySQLStr = MySQLStr & "ActDate, "
        MySQLStr = MySQLStr & "Curr, "
        MySQLStr = MySQLStr & "OrderSum, "
        MySQLStr = MySQLStr & "MyPayment, "
        MySQLStr = MySQLStr & "Avance1Type, "
        MySQLStr = MySQLStr & "Avance2Type, "
        MySQLStr = MySQLStr & "Comments) "
        MySQLStr = MySQLStr & "VALUES (N'" & UserName & "', "
        MySQLStr = MySQLStr & "N'" & OrderID & "', "
        MySQLStr = MySQLStr & "GETDATE ( ), "
        MySQLStr = MySQLStr & CurrCode & ", "
        MySQLStr = MySQLStr & Replace(CStr(OrderSum), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(MyPayment), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Avance1Type), ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(CStr(Avance2Type), ",", ".") & ", N'"
        MySQLStr = MySQLStr & Declarations.MyReason & "')"
        InitMyConn(True)
        InitMyRec(True, MySQLStr)
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 8")
    End Function


    Public Function CheckSalesPrices(ByVal OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �� ���������� �� ���� ������ � ������� �������� � 1 ���
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        If My.Settings.UseOffice = "LibreOffice" Then
            Return CheckSalesPricesLO(OrderNum)
        Else
            Return CheckSalesPricesExcel(OrderNum)
        End If

    End Function

    Public Function CheckSalesPricesExcel(ByRef OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �� ���������� �� ���� ������ � ������� �������� � 1 ��� (����� � Excel)
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)
        MySQLStr = "EXEC spp_ShippingCheck_CheckSalesPrice "
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckSalesPricesExcel = True
            trycloseMyRec()
        Else
            '---���� ������������� �����, �� � �������
            MyObj = CreateObject("Excel.Application")
            MyObj.SheetsInNewWorkbook = 1
            MyWRKBook = MyObj.Workbooks.Add

            MyWRKBook.ActiveSheet.Columns("A:C").ColumnWidth = 35

            '---��������� �����
            MyWRKBook.ActiveSheet.Range("B1") = "����� �� ������� ����� " & OrderNum
            MyWRKBook.ActiveSheet.Range("A4") = "� ���� ������ ����� �������� � 1 ��� ���������� ����:"

            MyWRKBook.ActiveSheet.Range("B1").Select()
            MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
            MyWRKBook.ActiveSheet.Range("A4").Select()
            MyWRKBook.ActiveSheet.Range("A4").Font.Bold = True

            MyWRKBook.ActiveSheet.Range("A5") = "��� ������"
            MyWRKBook.ActiveSheet.Range("B5") = "������� ���� �� �������"
            MyWRKBook.ActiveSheet.Range("C5") = "���� ��� �������� � 1 ��� �� �������"
            MyWRKBook.ActiveSheet.Range("A5:C5").Select()
            MyWRKBook.ActiveSheet.Range("A5:C5").Borders(5).LineStyle = -4142
            MyWRKBook.ActiveSheet.Range("A5:C5").Borders(6).LineStyle = -4142
            With MyWRKBook.ActiveSheet.Range("A5:C5").Borders(7)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:C5").Borders(8)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:C5").Borders(9)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:C5").Borders(10)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:C5").Borders(11)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("A5:C5").Interior
                .ColorIndex = 36
                .Pattern = 1
                .PatternColorIndex = -4105
            End With

            '---����� ����� ������
            i = 6
            Declarations.MyRec.MoveFirst()
            While Not Declarations.MyRec.EOF
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" + CStr(i)) = Declarations.MyRec.Fields("StockCode").Value
                MyWRKBook.ActiveSheet.Range("B" + CStr(i)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("C" + CStr(i)) = Declarations.MyRec.Fields("ConvertPrice").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            MyWRKBook.ActiveSheet.Range("A1").Select()
            MyObj.Application.Visible = True
            MyObj = Nothing
            CheckSalesPricesExcel = False
            trycloseMyRec()
        End If
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 11")
    End Function

    Public Function CheckSalesPricesLO(ByRef OrderNum As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - �� ���������� �� ���� ������ � ������� �������� � 1 ��� (����� � LibreOffice)
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim i As Integer

        OrderNum = Right("0000000000" & Trim(OrderNum), 10)
        MySQLStr = "EXEC spp_ShippingCheck_CheckSalesPrice "
        MySQLStr = MySQLStr & "N'" & OrderNum & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckSalesPricesLO = True
            trycloseMyRec()
        Else
            '---���� ������������� �����, �� � �������
            LOSetNotation(0)
            oServiceManager = CreateObject("com.sun.star.ServiceManager")
            oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
            oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
            Dim arg(1)
            arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
            arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
            oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
            oSheet = oWorkBook.getSheets().getByIndex(0)
            oFrame = oWorkBook.getCurrentController.getFrame
            '---������ �������
            oSheet.getColumns().getByName("A").Width = 4000
            oSheet.getColumns().getByName("B").Width = 4000
            oSheet.getColumns().getByName("C").Width = 4000
            '---��������� �����
            oSheet.getCellRangeByName("B1").String = "����� �� ������� ����� " & OrderNum
            oSheet.getCellRangeByName("A4").String = "� ���� ������ ����� �������� � 1 ��� ���������� ����:"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:B4", "Tahoma")
            LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1:B4")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1:B4", 11)
            '---��������� �������
            oSheet.getCellRangeByName("A5").String = "��� ������"
            oSheet.getCellRangeByName("B5").String = "������� ���� �� �������"
            oSheet.getCellRangeByName("C5").String = "���� ��� �������� � 1 ��� �� �������"
            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5:C5", "Calibri")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5:C5", 9)
            LOWrapText(oServiceManager, oDispatcher, oFrame, "A5:C5")
            oSheet.getCellRangeByName("A5:C5").CellBackColor = RGB(114, 251, 238)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
            Dim LineFormat As Object
            LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
            LineFormat.LineStyle = 0
            LineFormat.LineWidth = 70
            oSheet.getCellRangeByName("A5:C5").TopBorder = LineFormat
            oSheet.getCellRangeByName("A5:C5").RightBorder = LineFormat
            oSheet.getCellRangeByName("A5:C5").LeftBorder = LineFormat
            oSheet.getCellRangeByName("A5:C5").BottomBorder = LineFormat
            oSheet.getCellRangeByName("A5:C5").VertJustify = 2
            oSheet.getCellRangeByName("A5:C5").HoriJustify = 2
            '---����� ����� ������
            i = 6
            Declarations.MyRec.MoveFirst()
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("StockCode").Value
                oSheet.getCellRangeByName("B" & CStr(i)).Value = Declarations.MyRec.Fields("Price").Value
                oSheet.getCellRangeByName("C" & CStr(i)).Value = Declarations.MyRec.Fields("ConvertPrice").Value
                LOFormatCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":C" & CStr(i), 4)

                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            CheckSalesPricesLO = False
            trycloseMyRec()
            '----� ������ �����
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "$A$1"
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
            '----���������
            oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
            oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
        End If
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 12")
    End Function
End Module
