Module CHKFunctions

    Public Function CheckDataInProposal(ByVal MyOrder As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ���������� ������ ����������� - �������� ������������ ��� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   '������� ������
        Dim cmd As New ADODB.Command
        Dim MyCustomerCode

        '----���������� ���� � Scala
        MySQLStr = "SELECT ISNULL(SL010300.SL01001, N'') AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON tbl_OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE     (tbl_OR010300.OR01001 = N'" & MyOrder & "') "
        MySQLStr = MySQLStr & "GROUP BY SL010300.SL01001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            MyCustomerCode = Declarations.MyRec.Fields("CC").Value
            If Trim(MyCustomerCode) <> "" Then
                CheckDataInProposal = True
            Else
                MsgBox("���������� ������� ������ ���������� � Scala. ��� �������� ����������� � ���� 0 ���� ���������� ������� ������� ���������� � Scala ��� � ����������� � ������� �������� ����������.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '-----�������� ������ ����� ��� �� ���� �����, ��� � ����������
        Dim MyRezStr As String
        MyRezStr = CheckSalesman(Declarations.SalesmanCode, MyCustomerCode)
        If MyRezStr <> "" Then
            MsgBox(MyRezStr, MsgBoxStyle.OkOnly, "��������!")
            CheckDataInProposal = False
            Exit Function
        End If


        '----� ������ ��� ������
        MySQLStr = "SELECT COUNT(OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("� ������ ����������� ��� �� ����� ������. ��� �������� ����������� � ���� 0 ���� ���������� ������� ���� �� 1 ������.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----��� ������ ���� � Scala
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(SC010300.SC01001 IS NULL)"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("� ������ ������ ���� ������, ������� ���������� � Scala. ��� �������� ����������� � ���� 0 ���� ���������� ������� ������ ������ � Scala ��� � ����������� � ������� �������� ������ �� ������������ � Scala.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----������ � ������ �� �������� ����������������
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(tbl_ItemCard0300.IsBlocked = N'1') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("� ������ ������ ���� ��������������� ������. ��� �������� ����������� � ���� 0 ���� ���������� ������� ��� ������ �� �����������, ��� � ����������� � ������� �������� ������ �� ����������������� � Scala.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----� ����� �� �������� ��������� ������
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(SC010300.SC01066 = 8) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("� ������ ������ ���� ��������� ������. ������������� �� � ����������� � ������� ���������. ��� �������� ����������� � ���� 0 ���� ���������� ������� ��� ������ �� �����������.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----��������, ��� ���������� ���������� ���� (������ �������)
        MySQLStr = "SELECT PriceCond, ReadyDate, DeliveryDate, ExpirationDate "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            trycloseMyRec()
            Exit Function
        Else
            'If Declarations.MyRec.Fields("PriceCond").Value = "�������� �� �����" Then
            'If Declarations.MyRec.Fields("DeliveryDate").Value < Now() Then
            'MsgBox("� ������ ������ ���� �������� ������ ������� ����. ��������������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
            'CheckDataInProposal = False
            'trycloseMyRec()
            'Exit Function
            'Else
            'CheckDataInProposal = True
            'End If
            'Else
            'If Declarations.MyRec.Fields("ReadyDate").Value < Now() Then
            '    MsgBox("� ������ ������ ���� ���������� � �������� ������ ������� ����. ��������������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
            '    CheckDataInProposal = False
            '    trycloseMyRec()
            '    Exit Function
            'Else
            '    CheckDataInProposal = True
            'End If
            'End If
            If Declarations.MyRec.Fields("ExpirationDate").Value < Now() Then
                MsgBox("� ������ ������ ���� �������� ����������� ������ ������� ����. ��������������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                trycloseMyRec()
                Exit Function
            Else
                CheckDataInProposal = True
            End If
        End If
        trycloseMyRec()

        '----��������, ��� ����� ���������� ����� ��������
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM  tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(WeekQTY IS NULL) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("������ �������� ����������� � ����� 0 ����. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("� ������ ������ ���� ������ � ���������������� ������� ��������. ��� �������� ����������� � ���� 0 ���� ���������� �� ���� ������� ���������� ����� ��������.", MsgBoxStyle.Critical, "��������!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '-----��������, ��� ������� ��������� � �� � � Scala �����
        cmd.ActiveConnection = Declarations.MyConn
        cmd.CommandText = "spp_Agents_CheckCPData"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.CommandTimeout = 300

        Dim MyParam As ADODB.Parameter                  '������������ �������� ����� 1
        Dim MyParam1 As ADODB.Parameter                 '������������ �������� ����� 2

        Dim MyRSTR As String
        MyRSTR = ""

        MyParam = cmd.CreateParameter("@MyCPID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
        cmd.Parameters.Append(MyParam)
        MyParam.Value = MyOrder

        MyParam1 = cmd.CreateParameter("@MyRetValue", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
        cmd.Parameters.Append(MyParam1)
        MyParam1.Value = ""

        Try
            cmd.Execute()
            MyRSTR = MyParam1.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        End Try

        If Trim(MyRSTR) <> "" Then
            MsgBox("� ������������ ����������� ���� ������, ������� ��������� ������� ���������� �� �������� � �������� ������: " & Chr(13) & Chr(10) & Trim(MyRSTR), MsgBoxStyle.Critical, "��������!")
            CheckDataInProposal = False
            Exit Function
        End If


        CheckDataInProposal = True
    End Function
End Module
