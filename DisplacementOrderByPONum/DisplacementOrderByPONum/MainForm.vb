Public Class MainForm

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyDs As New DataSet                       '

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            Declarations.ScalaDate = CDate(Scala.ActiveFrame.Parent.ScalaDate)


            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("�� ������ ��� ��������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Application.Exit()
            Else
                Declarations.SalesmanCode = Declarations.MyRec.Fields("SC").Value
                Declarations.SalesmanName = Declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
            End If
        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try

        '---���������� ����� (ComboBox)
        BuildWHList()

        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today

    End Sub

    Private Sub BuildWHList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � Combobox ������ ������� � ����� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "SC23002" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "SC23001"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� �����������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckEmpty() = True Then
            If CheckCorrect(TextBox1.Text, ComboBox1.SelectedValue) = True Then
                CreateDisplacementOrder(TextBox1.Text, ComboBox1.SelectedValue)
            End If
        End If
    End Sub

    Private Function CheckEmpty() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ���� ����������� �����
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        '---����� ������ �� �������
        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""����� ������ �� �������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
            TextBox1.Select()
            CheckEmpty = False
            Exit Function
        End If

        '---���� ������
        If DateTimePicker1.Value = Today And DateTimePicker2.Value = Today Then
            MsgBox("������� ���� �������������� �������� � ������� ������ (�� ������ ���� ����� ����� ����� � �������������� ���� ��������� ������ �� ������ ���� ����� ������������ ���)", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckEmpty = False
            Exit Function
        End If

        CheckEmpty = True
    End Function

    Private Function CheckCorrect(ByVal MyOrderNum As String, ByVal MyWhNum As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ��������� ��������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim SOWhNum As String                         '����� ������ ������ �� �������
        Dim POWhNum As String                         '����� ������ ������ �� �������

        MyOrderNum = Microsoft.VisualBasic.Right("0000000000" & Trim(MyOrderNum), 10)
        MyWhNum = Trim(MyWhNum)

        '---������������� ������ �� �����������
        MySQLStr = "SELECT COUNT(PC01001) AS �� "
        MySQLStr = MySQLStr & "FROM PC010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrderNum & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("��").Value = 0 Then
            trycloseMyRec()
            MsgBox("������ � ������� " & MyOrderNum & " ��� � ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckCorrect = False
            Exit Function
        End If
        trycloseMyRec()

        '---� ������ �� ������� ���������� ����� �� ������� � ��� ����� ���������� �� ������ ����������
        MySQLStr = "SELECT View_2.Expr1 AS CC "
        MySQLStr = MySQLStr & "FROM PC010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR01001, MAX(OR01050) AS Expr1 "
        MySQLStr = MySQLStr & "FROM (SELECT OR01001, OR01050 "
        MySQLStr = MySQLStr & "FROM OR010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT OR20001, OR20050 "
        MySQLStr = MySQLStr & "FROM OR200300 WITH(NOLOCK)) AS View_1 "
        MySQLStr = MySQLStr & "GROUP BY OR01001) AS View_2 ON PC010300.PC01060 = View_2.OR01001 "
        MySQLStr = MySQLStr & "WHERE (PC010300.PC01001 = N'" & MyOrderNum & "') "
        MySQLStr = MySQLStr & "GROUP BY View_2.Expr1 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            SOWhNum = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
            If SOWhNum = MyWhNum Then
            Else
                MsgBox("� ������ �� ������� �  ������� " & MyOrderNum & " ������ ����� �� ������� �� ������ " & SOWhNum & ". ���� ����� ������ ���������� �� ����� " & MyWhNum & ".", MsgBoxStyle.Critical, "��������!")
                CheckCorrect = False
                Exit Function
            End If

        End If


        '---������ ������� ������ �� ������� � ���������� ���������
        MySQLStr = "SELECT PC01023 AS CC "
        MySQLStr = MySQLStr & "FROM PC010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PC01001 = N'" & MyOrderNum & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("������ � ������� " & MyOrderNum & " ��� � ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckCorrect = False
            Exit Function
        Else
            POWhNum = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
            If POWhNum = MyWhNum Then
                MsgBox("� ������ �� ������� �  ������� " & MyOrderNum & " ������ ����� " & POWhNum & ". �� ��������� ����������� ��� �� ���� �� �����, ����� ������ ������. ", MsgBoxStyle.Critical, "��������!")
                CheckCorrect = False
                Exit Function
            End If
        End If

        '---������ ���������� (��� ��������������� ��� ���������)
        MySQLStr = "SELECT COUNT(PC010300.PC01001) AS CC "
        MySQLStr = MySQLStr & "FROM SC330300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN "
        MySQLStr = MySQLStr & "PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND PC190300.PC19002 = PC030300.PC03002 INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT UPPER(SYPD001) AS SYPD001, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_1 ON UPPER(PC010300.PC01046) = View_1.SYPD001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON SC330300.SC33001 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND "
        'MySQLStr = MySQLStr & "(LTRIM(RTRIM(PC010300.PC01060)) <> N'') AND "
        MySQLStr = MySQLStr & "(PC010300.PC01001 = N'" & MyOrderNum & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            trycloseMyRec()
            MsgBox("�� ������ �� ������� � ������� " & MyOrderNum & " ���������� ������. ��������� � ����������� ������� ���. ��� ����� ��� �� ������, ��� ��� �������� �� ����� ������ ���������������, ��� ��� ������� � ���������. ", MsgBoxStyle.Critical, "��������!")
            CheckCorrect = False
            Exit Function
        End If
        trycloseMyRec()

        '---����� - �������� � ����� ���������� �������� �������� ������������� �����
        If IsRawMaterialsWH(POWhNum) = True And IsRawMaterialsWH(MyWhNum) = True Then
            MsgBox("����� ������ �� ������� � ������� " & MyOrderNum & " � ����� ���������� " & MyWhNum & " �������� �������� ������������� �����. ������ ����� �� ����������� � ������ ������ ������������� ����� �� ������ ������.", MsgBoxStyle.Critical, "��������!")
            CheckCorrect = False
            Exit Function
        End If

        CheckCorrect = True
    End Function

    Private Function CreateDisplacementOrder(ByVal MyOrderNum As String, ByVal MyWhNum As String) As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� �����������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim cmd As New ADODB.Command                    '������� (spp ���������)
        Dim MyParam As ADODB.Parameter                  '������������ �������� ����� 1
        Dim MyParam1 As ADODB.Parameter                 '������������ �������� ����� 2
        Dim MyParam2 As ADODB.Parameter                 '������������ �������� ����� 3
        Dim MyParam3 As ADODB.Parameter                 '������������ �������� ����� 4
        Dim MyParamR As ADODB.Parameter                 '������������ �������� ����� 5 (���� ������������)
        Dim MyRezStr As String                          '������ � ������������� ����������

        MyOrderNum = Microsoft.VisualBasic.Right("0000000000" & Trim(MyOrderNum), 10)
        MyWhNum = Trim(MyWhNum)
        MyRezStr = ""
        InitMyConn(False)
        Try
            cmd.ActiveConnection = Declarations.MyConn
            cmd.CommandText = "spp_DisplacementOrderCreation"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.CommandTimeout = 300

            '----�������� ����������---------------------------------------------------
            '---����� ������ �� �������
            MyParam = cmd.CreateParameter("@MyPurchaseOrderNum", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam)
            '---����� ����������
            MyParam1 = cmd.CreateParameter("@MyDestWarNo", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam1)
            '---���� ��������
            MyParam2 = cmd.CreateParameter("@MyOrderDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam2)
            '---���� ���������
            MyParam3 = cmd.CreateParameter("@MyShipDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam3)
            '---������������ �������� (������)
            MyParamR = cmd.CreateParameter("@MyRezStr", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
            cmd.Parameters.Append(MyParamR)

            '----�������� ����������---------------------------------------------------
            '---����� ������ �� �������
            MyParam.Value = Microsoft.VisualBasic.Right("0000000000" + Trim(MyOrderNum), 10)
            '---����� ����������
            MyParam1.Value = Trim(MyWhNum)
            '---���� ��������
            MyParam2.Value = DateTimePicker1.Value
            '---���� ���������
            MyParam3.Value = DateTimePicker2.Value
            '---������ �������� ���������------------------------------------------------
            '---����������
            '--SetBlock(MyParam.Value) --���������� � �������� ���������

            '---���������
            cmd.Execute()
            MyRezStr = MyRezStr + LTrim(RTrim(MyParamR.Value))
            '---������ ����������
            '--RemoveBlock()


        Catch ex As Exception
            MyRezStr = MyRezStr + ex.Message
        End Try

        MsgBox("��������� �������� ������ �� ����������� ���������. " & MyRezStr, MsgBoxStyle.OkOnly, "��������!")
    End Function
End Class
