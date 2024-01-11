Public Class EditHeader

    Public StartParam As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----���������� �����������
            If SaveHeader() = True Then
                If CheckEmptyHDR() = True Then
                    Me.Close()
                End If
            End If
        End If
    End Sub

    Private Sub EditHeader_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        'If ComboBox2.Text = "0" Then
        '    Label9.Text = "RUR"
        '    Label18.Text = CStr(GetExchangeRate(0, Now()))
        'ElseIf ComboBox2.Text = "1" Then
        '    Label9.Text = "USD"
        '    Label18.Text = CStr(GetExchangeRate(1, Now()))
        'ElseIf ComboBox2.Text = "6" Then
        '    Label9.Text = "CNY"
        '    Label18.Text = CStr(GetExchangeRate(6, Now()))
        'Else
        '    Label9.Text = "EUR"
        '    Label18.Text = CStr(GetExchangeRate(12, Now()))
        'End If
        Label9.Text = ComboBox2.Text
        Label18.Text = CStr(GetExchangeRate(ComboBox2.SelectedValue, Now()))
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ����� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        'If ComboBox3.Text = "0" Then
        '    Label10.Text = "RUR"
        '    Label19.Text = CStr(GetExchangeRate(0, Now()))
        'ElseIf ComboBox3.Text = "1" Then
        '    Label10.Text = "USD"
        '    Label19.Text = CStr(GetExchangeRate(1, Now()))
        'Else
        '    Label10.Text = "EUR"
        '    Label19.Text = CStr(GetExchangeRate(12, Now()))
        'End If
        Label10.Text = ComboBox3.Text
        Label19.Text = CStr(GetExchangeRate(ComboBox3.SelectedValue, Now()))
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox2, True, True, True, False)
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox3, True, True, True, False)
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox4.Text = "��������� �� ������" Then
            Declarations.IsSelfDelivery = 1
        Else
            Declarations.IsSelfDelivery = 0
        End If
        Me.SelectNextControl(ComboBox4, True, True, True, False)
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub EditHeader_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    '
        Dim MyAdapter3 As SqlClient.SqlDataAdapter    '
        Dim MyAdapter4 As SqlClient.SqlDataAdapter    '
        Dim MyDs As New DataSet                       '
        Dim MyDs1 As New DataSet                      '
        Dim MyDs2 As New DataSet                      '
        Dim MyDs3 As New DataSet                      '
        Dim MyDs4 As New DataSet                      '
        Dim MyPRID As String                          'ID �����������
        Dim MyCCode As String                         '��� ����������
        Dim MyCName As String                         '��� ����������
        Dim MyCAddr As String                         '����� ����������
        Dim MyWHNum As String                         '����� ������
        Dim MyDocCode As Integer                      '��� ���������
        Dim MyPRCurrCode As Integer                   '��� ������ �����������
        Dim MyInvCurrCode As Integer                  '��� ������ ��
        Dim MyComment As String                       '����������
        Dim MyPriceCond As String                     '������� ����������� ����
        Dim MyReadyDate As DateTime                   '���� ���������� �� ������
        Dim MyDeliveryAddr As String                  '����� ��������
        Dim MyDeliveryDate As DateTime                '���� ��������
        Dim MyPaymentCond As String                   '������� �������
        Dim MyExpirationDate As DateTime              '����, �� ������� ��������� �����������
        Dim MyPartialShipment As Integer              '����������� ��������� �������� (0 - ���, 1 - ��)
        Dim MyAgentName As String                     '��� ��������� ������, ���������� ������������ �����������
        Dim MyCPState As String                       '��������� ������������� �����������


        '----���������� ������ ������� � ComboBox
        MySQLStr = "SELECT SC23001, SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT '' AS SC23001,'' AS SC23002 "
        MySQLStr = MySQLStr & "ORDER BY SC23001 "
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

        '----���������� ������ ����� ���������� � ComboBox
        MySQLStr = "SELECT SY24002, SY24003 "
        MySQLStr = MySQLStr & "FROM SY240300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SY24001 = N'DC') "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT '' AS SY24002,'' AS SY24003 "
        MySQLStr = MySQLStr & "ORDER BY SY24002 "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.Fill(MyDs1)
            ComboBox6.DisplayMember = "SY24003" '��� �� ��� ����� ������������
            ComboBox6.ValueMember = "SY24002"   '��� �� ��� ����� ���������
            ComboBox6.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----���������� ������ ����� � ComboBox ������ ���������
        MySQLStr = "SELECT SYCD001, SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD009 <> N'') "
        MySQLStr = MySQLStr & "AND (SYCD009 NOT IN ('FIM', 'FRF', 'SEK', 'DK', 'DM', 'FI1', 'ROL')) "
        Try
            MyAdapter3 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter3.Fill(MyDs3)
            ComboBox2.DisplayMember = "SYCD009" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "SYCD001"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs3.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----���������� ������ ����� � ComboBox ������ ��
        MySQLStr = "SELECT SYCD001, SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD009 <> N'') "
        MySQLStr = MySQLStr & "AND (SYCD009 = N'RUB') "
        MySQLStr = MySQLStr & "AND (SYCD009 NOT IN ('FIM', 'FRF', 'SEK', 'DK', 'DM', 'FI1', 'ROL')) "
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.Fill(MyDs4)
            ComboBox3.DisplayMember = "SYCD009" '��� �� ��� ����� ������������
            ComboBox3.ValueMember = "SYCD001"   '��� �� ��� ����� ���������
            ComboBox3.DataSource = MyDs4.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----���������� ������ ��������� �� � ComboBox
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_SalesmanWorkPlace4_CPState "
        MySQLStr = MySQLStr & "ORDER BY Name "
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.Fill(MyDs2)
            ComboBox7.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox7.ValueMember = "Name"   '��� �� ��� ����� ���������
            ComboBox7.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        If StartParam = "Create" Then
            '----��������� ������ �����������
            Label3.Text = Microsoft.VisualBasic.Right("0000000000" & CStr(GetNewID()), 10)
            ComboBox4.SelectedItem = "��������� �� ������"

        Else '----�������� �� �������������� - ���������� ������������ �����
            '----��������� ������ �����������
            MyPRID = Form1.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString

            MySQLStr = "SELECT  tbl_OR010300.OR01001, "
            MySQLStr = MySQLStr & "tbl_OR010300.OR01003, "
            MySQLStr = MySQLStr & " ISNULL(SL010300.SL01001, N'') AS SL01001, "
            MySQLStr = MySQLStr & "ISNULL(SL010300.SL01002, N'') AS SL01002, "
            MySQLStr = MySQLStr & "LTRIM(RTRIM(ISNULL(SL010300.SL01003, N'') + ' ' + ISNULL(SL010300.SL01004, N'') + ' ' + ISNULL(SL010300.SL01005, N''))) AS SL01003, "
            MySQLStr = MySQLStr & "tbl_OR010300.OR01050, "
            MySQLStr = MySQLStr & "tbl_OR010300.OR01028, "
            MySQLStr = MySQLStr & "tbl_OR010300.OR01116, "
            MySQLStr = MySQLStr & "LTRIM(RTRIM(ISNULL(View_1.OR17005, N'') + ' ' + ISNULL(View_1.OR17006, N''))) AS OR17005, "
            MySQLStr = MySQLStr & "tbl_OR010300.CName, "
            MySQLStr = MySQLStr & "tbl_OR010300.CAddr, "
            MySQLStr = MySQLStr & "tbl_OR010300.PriceCond, "
            MySQLStr = MySQLStr & "tbl_OR010300.ReadyDate, "
            MySQLStr = MySQLStr & "tbl_OR010300.DeliveryAddr, "
            MySQLStr = MySQLStr & "tbl_OR010300.DeliveryDate, "
            MySQLStr = MySQLStr & "tbl_OR010300.PaymentCond, "
            MySQLStr = MySQLStr & "tbl_OR010300.ExpirationDate, "
            MySQLStr = MySQLStr & "tbl_OR010300.PartialShipment, "
            MySQLStr = MySQLStr & "ISNULL(tbl_OR010300.OR01065, SL010300.SL01085) AS DocCode, "
            MySQLStr = MySQLStr & "ISNULL(tbl_OR010300.AgentName, '') AS AgentName, "
            MySQLStr = MySQLStr & "ISNULL(tbl_OR010300.CPState, '') AS CPState "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT * "
            MySQLStr = MySQLStr & "FROM tbl_OR170300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR17002 = N'000000') AND "
            MySQLStr = MySQLStr & "(OR17003 = N'000000') AND "
            MySQLStr = MySQLStr & "(OR17004 = N'510')) AS View_1 ON "
            MySQLStr = MySQLStr & "tbl_OR010300.OR01001 = View_1.OR17001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SL010300 ON tbl_OR010300.OR01003 = SL010300.SL01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01001 = N'" & MyPRID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
            Else
                '----����� ����������� (��� � MyPRID) 
                '----��� ����������
                If Trim(Declarations.MyRec.Fields("SL01001").Value.ToString) = "" Then
                    MyCCode = Declarations.MyRec.Fields("OR01003").Value
                    TextBox2.ReadOnly = False
                    TextBox3.ReadOnly = False
                Else
                    MyCCode = Declarations.MyRec.Fields("SL01001").Value
                    TextBox2.ReadOnly = True
                    TextBox3.ReadOnly = True
                End If
                '----��� ����������
                If Declarations.MyRec.Fields("SL01002").Value.ToString = "" Then
                    MyCName = Declarations.MyRec.Fields("CName").Value
                Else
                    MyCName = Declarations.MyRec.Fields("SL01002").Value
                End If
                '----����� ����������
                If Declarations.MyRec.Fields("SL01003").Value.ToString = "" Then
                    MyCAddr = Declarations.MyRec.Fields("CAddr").Value
                Else
                    MyCAddr = Declarations.MyRec.Fields("SL01003").Value
                End If
                '----����� ������
                MyWHNum = Declarations.MyRec.Fields("OR01050").Value
                '----��� ���������
                MyDocCode = Declarations.MyRec.Fields("DocCode").Value
                '----������ �����������
                MyPRCurrCode = Declarations.MyRec.Fields("OR01028").Value
                '----������ ��
                MyInvCurrCode = Declarations.MyRec.Fields("OR01116").Value
                '----����������
                MyComment = Declarations.MyRec.Fields("OR17005").Value
                '----������� ����������� ����
                MyPriceCond = Declarations.MyRec.Fields("PriceCond").Value
                '----������� �������
                MyPaymentCond = Declarations.MyRec.Fields("PaymentCond").Value
                '----����, �� ������� ��������� �����������
                MyExpirationDate = Declarations.MyRec.Fields("ExpirationDate").Value
                '----����������� ��������� �������� (0 - ���, 1 - ��)
                MyPartialShipment = Declarations.MyRec.Fields("PartialShipment").Value
                '----��� ��������� ������
                MyAgentName = Declarations.MyRec.Fields("AgentName").Value
                '----��������� ��
                MyCPState = Declarations.MyRec.Fields("CPState").Value
                trycloseMyRec()

                '----����� �����������
                Label3.Text = MyPRID
                '----��� ����������
                TextBox1.Text = MyCCode
                '----��� ����������
                TextBox2.Text = MyCName
                '----����� ����������
                TextBox3.Text = MyCAddr
                '----����� ������
                ComboBox1.SelectedValue = MyWHNum
                '----��� ���������
                ComboBox6.SelectedValue = MyDocCode
                '----������ �����������
                ComboBox2.SelectedValue = MyPRCurrCode
                '----������ ��
                ComboBox3.SelectedValue = MyInvCurrCode
                '----����������
                TextBox4.Text = MyComment
                '----������� ����������� ����
                ComboBox4.Text = MyPriceCond
                If ComboBox4.Text = "��������� �� ������" Then
                    Declarations.IsSelfDelivery = 1
                Else
                    Declarations.IsSelfDelivery = 0
                End If
                '----������� �������
                TextBox8.Text = MyPaymentCond
                '----����, �� ������� ��������� �����������
                DateTimePicker3.Value = MyExpirationDate
                '----����������� ��������� �������� (0 - ���, 1 - ��)
                If MyPartialShipment = 0 Then
                    ComboBox5.Text = "���"
                Else
                    ComboBox5.Text = "��"
                End If
                '----��� ��������� ������
                TextBox5.Text = MyAgentName
                '----��������� ��
                ComboBox7.SelectedValue = MyCPState
            End If
        End If

        '----����� � ������ ����
        TextBox1.Select()
        Declarations.MyOrderNum = Trim(Label3.Text)
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerSelect = New CustomerSelect
        MyCustomerSelect.StartParam = "CP"
        MyCustomerSelect.ShowDialog()
        CheckCustomerBlock()
    End Sub


    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ���� �� ��� � Scala, ���� �� - �� ��������� ��� � ����� 
        '// ����� ��������� - �� ������������ �� ������ � Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As Double
        Dim MyRezStr As String

        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyRez = Declarations.MyRec.Fields("CC").Value
        trycloseMyRec()
        If MyRez = 1 Then
            MyRezStr = CheckSalesman(Declarations.SalesmanCode, Trim(TextBox1.Text))
            If MyRezStr = "" Then
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                MySQLStr = "SELECT SL01002, SL01003 + ' ' + SL01004 + ' ' + SL01005 AS SL01003, "
                MySQLStr = MySQLStr & "SL01085, SL01098 "
                MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                TextBox2.Text = Declarations.MyRec.Fields("SL01002").Value
                TextBox3.Text = Declarations.MyRec.Fields("SL01003").Value
                ComboBox6.SelectedValue = Declarations.MyRec.Fields("SL01085").Value
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("SL01098").Value
                trycloseMyRec()
                CheckCustomerBlock()
            Else
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = False
                MsgBox(MyRezStr, MsgBoxStyle.OkOnly, "��������!")
            End If
        Else
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
        End If
    End Sub

    Private Sub CheckCustomerBlock()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � ���������� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustomerIsCredit As Integer

        '---�������� ���������� �� ������� � ����������
        MySQLStr = "SELECT tbl_CustomerCard0300.IsBlocked, tbl_CustomerCard0300.DataFrom, tbl_CustomerCard0300.DataTo, CASE WHEN (SL01024 = N'0' OR "
        MySQLStr = MySQLStr & "SL01024 = N'00') AND SL01037 = 0 THEN 0 ELSE 1 END AS IsCredit "
        MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON SL010300.SL01001 = tbl_CustomerCard0300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(TextBox1.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("IsCredit").Value = 1 Then   '---��������� ������
                CustomerIsCredit = 1
            Else
                CustomerIsCredit = 0
            End If
            If Declarations.MyRec.Fields("IsBlocked").Value = 1 Then
                If Declarations.MyRec.Fields("IsCredit").Value = 1 Then   '---��������� ������
                    trycloseMyRec()
                    MsgBox("������ � ����� " & Trim(TextBox1.Text) & " �������� ��������� �������� � � ��������� ������ ������������. ���������� ��� ���������� �������� ��������� ����� ���������� �������� � ���������� � ������� 2 - � ��� �������� �� ������� �������. ��� ������ ������� ������� ������ � 1 ��� ��� ������ ���������� �� �������� ����� �������� ������ ����� ��������� � �������� ������� ���������� �� ������ ���������� ��������, ����� ���������� (��� �������������) � ��� �������������. " & _
                        "��� ����� ����� ���������� ���������� ������� ������ '�������� �������' �� �������. ��� ����� ���������� �� ������ ���������� �������� ���������� ������� ������ '���������� ���������� ��������' �� �������. ", vbOKOnly, "��������!")
                Else                                                      '---����������� ������
                    trycloseMyRec()
                    MsgBox("������ � ����� " & Trim(TextBox1.Text) & " �������� ����������� �������� � �� ���� � ������� 2 - � ��� �� ���� ��������. ��� ������ ������� ����� ��������� ������������� ����������� ��� ������ ���������, ��� ��������� ������� �� ����������. " & _
                        "��� ����� ����� ���������� � ������ �� ��������� ���������� ������� ������ '�������� �������' �� �������. ", vbOKOnly, "��������!")
                End If
            Else
                If CustomerIsCredit = 1 Then                    '---��������� ������ ��� ��������� �����������
                    MySQLStr = "SELECT DataFrom, DataTo "
                    MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 "
                    MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        MsgBox("������ Salesman Workplace 4 ������� CheckCustomerBlock �������� - �� ����� �� ���� ��������. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
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
                                MsgBox("��������! ���������� ��� ��� �� ������� ���� �������� �������� �������� � ������ ��������. ��������� ����� ������� � �������� ������ � ��� � ����.", vbCritical, "��������!")
                                trycloseMyRec()
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �������������� ���� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----���������� �����������
            If SaveHeader() = True Then
                '----�������� ���� ��������������
                MyOrderLines = New OrderLines
                MyOrderLines.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckEmptyHDR() = True Then
            Me.Close()
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""��� ����������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
            TextBox1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If TextBox2.ReadOnly = False And Trim(TextBox2.Text) = "" Then
            MsgBox("���� ""����������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
            TextBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If TextBox3.ReadOnly = False And Trim(TextBox3.Text) = "" Then
            MsgBox("���� ""����� ����������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
            TextBox3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = "" Then
            MsgBox("����� �������� ������ ���� ������", MsgBoxStyle.Critical, "��������!")
            ComboBox1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If ComboBox6.SelectedValue = "" Then
            MsgBox("��� ��������� ������ ���� ������", MsgBoxStyle.Critical, "��������!")
            ComboBox6.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox2.Text) = "" Then
            MsgBox("���������� ������� ������ �����������", MsgBoxStyle.Critical, "��������!")
            ComboBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox3.Text) = "" Then
            MsgBox("���������� ������� ������ ����� - �������", MsgBoxStyle.Critical, "��������!")
            ComboBox3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox4.Text) = "" Then
            MsgBox("���������� ������� ������� ����������� ��� - ������� ��������� ��������.", MsgBoxStyle.Critical, "��������!")
            Button7.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(TextBox8.Text) = "" Then
            MsgBox("���� ""������� ������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������!")
            TextBox8.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If DateTimePicker3.Value < Now() Then
            MsgBox("���� ""���� �������� ����������� - ��:"" ������ ���� ������ �������", MsgBoxStyle.Critical, "��������!")
            DateTimePicker3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox5.Text) = "" Then
            MsgBox("���������� ������� - �������� �� ��������� ��������", MsgBoxStyle.Critical, "��������!")
            ComboBox5.Select()
            CheckFormFilling = False
            Exit Function
        End If

        CheckFormFilling = True
    End Function

    Private Function SaveHeader() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ ��������� � �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyPRID As String                          'ID �����������
        Dim MyCCode As String                         '��� ����������
        Dim MyCName As String                         '��� ����������
        Dim MyCAddr As String                         '����� ����������
        Dim MyWHNum As String                         '����� ������
        Dim MyDocCode As Integer                      '��� ���������
        Dim MyPRCurrCode As Integer                   '��� ������ �����������
        Dim MyInvCurrCode As Integer                  '��� ������ ��
        Dim MyComment As String                       '����������
        Dim MyPriceCond As String                     '������� ����������� ����
        Dim MyReadyDate As String                     '���� ���������� �� ������
        Dim MyDeliveryAddr As String                  '����� ��������
        Dim MyDeliveryDate As String                  '���� ��������
        Dim MyPaymentCond As String                   '������� �������
        Dim MyExpirationDate As String                '����, �� ������� ��������� �����������
        Dim MyPartialShipment As Integer              '����������� ��������� �������� (0 - ���, 1 - ��)
        Dim MySQLStr As String                        '������� ������
        Dim MyAgentName As String                     '��� ��������� ������
        Dim MyCPState As String                       '��������� ��

        '----����� �����������
        MyPRID = Trim(Label3.Text)
        '----��� ����������
        MyCCode = Trim(TextBox1.Text)
        '----��� ����������
        If TextBox2.Enabled = True Then
            MyCName = Trim(TextBox2.Text)
        Else
            MyCName = ""
        End If
        '----����� ����������
        If TextBox3.Enabled = True Then
            MyCAddr = Trim(TextBox3.Text)
        Else
            MyCAddr = ""
        End If
        '----����� ������
        MyWHNum = ComboBox1.SelectedValue
        '----��� ���������
        MyDocCode = ComboBox6.SelectedValue
        '----������ �����������
        MyPRCurrCode = Trim(ComboBox2.SelectedValue)
        '----������ ��
        MyInvCurrCode = Trim(ComboBox3.SelectedValue)
        '----����������
        MyComment = Trim(TextBox4.Text)
        '----������� ����������� ����
        MyPriceCond = Trim(ComboBox4.Text)
        '----���� ���������� �� ������
        MyReadyDate = "01/01/1900"
        MyDeliveryAddr = ""
        '----���� ��������
        MyDeliveryDate = "01/01/1900"
        '----������� �������
        MyPaymentCond = Trim(TextBox8.Text)
        '----����, �� ������� ��������� �����������
        MyExpirationDate = DatePart(DateInterval.Day, DateTimePicker3.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker3.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker3.Value)
        '----����������� ��������� �������� (0 - ���, 1 - ��)
        If ComboBox5.Text = "���" Then
            MyPartialShipment = 0
        Else
            MyPartialShipment = 1
        End If
        '----��� ��������� ������
        MyAgentName = TextBox5.Text
        '----��������� ��
        MyCPState = ComboBox7.SelectedValue

        Try
            MySQLStr = "EXEC spp_SalesWorkplace4_AddOrderHeader1 "
            MySQLStr = MySQLStr & "N'" & MyPRID & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyCCode, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyCName, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyCAddr, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & MyWHNum & "', "
            MySQLStr = MySQLStr & MyDocCode & ", "
            MySQLStr = MySQLStr & MyPRCurrCode & ", "
            MySQLStr = MySQLStr & MyInvCurrCode & ", "
            MySQLStr = MySQLStr & "N'" & Replace(MyComment, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyPriceCond, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & MyReadyDate & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyDeliveryAddr, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & MyDeliveryDate & "', "
            MySQLStr = MySQLStr & "N'" & Replace(MyPaymentCond, "'", "''") & "', "
            MySQLStr = MySQLStr & "N'" & Declarations.SalesmanCode & "', "
            MySQLStr = MySQLStr & "N'" & MyExpirationDate & "', "
            MySQLStr = MySQLStr & MyPartialShipment & ", "
            MySQLStr = MySQLStr & "N'" & MyAgentName & "', "
            MySQLStr = MySQLStr & "N'" & MyCPState & "' "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        Catch ex As Exception
            MsgBox(ex.ToString)
            SaveHeader = False
            Exit Function
        End Try


        SaveHeader = True
    End Function

    Private Function CheckEmptyHDR() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ���� �� ������ � �����������.
        '// ���� ��� - �������� ��������� � ����������
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyPRID As String                          'ID �����������
        Dim MySQLStr As String                        '������� ������
        Dim MyRez As Double                           '��������� - ���������� �����
        Dim MyRez1 As VariantType                     '��������� ������


        '----����� �����������
        MyPRID = Trim(Label3.Text)
        MySQLStr = "SELECT COUNT(*) AS CL "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyPRID & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyRez = Declarations.MyRec.Fields("CL").Value
        trycloseMyRec()

        If MyRez = 0 Then '---��� ����� � ������
            MyRez1 = MsgBox("� ����������� ����� " & MyPRID & " ��� �� ����� ������. ������� ���������? ", MsgBoxStyle.YesNo, "��������!")
            If MyRez1 = vbYes Then
                MySQLStr = "DELETE FROM tbl_OR010300 "
                MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyPRID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                MySQLStr = "DELETE FROM tbl_OR170300 "
                MySQLStr = MySQLStr & "WHERE (OR17001 = N'" & MyPRID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                CheckEmptyHDR = True
            Else
                CheckEmptyHDR = False
            End If
        Else
            CheckEmptyHDR = True
        End If
    End Function

    Private Sub ComboBox5_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox5, True, True, True, False)
    End Sub

    Private Sub DateTimePicker3_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.CloseUp
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker3, True, True, True, False)
    End Sub

    Private Sub DateTimePicker3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �� ��������� ���� �� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ����� ������ �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim MySQLStr As String

        If CheckFormFilling() = True Then
            '----���������� �����������
            If SaveHeader() = True Then
                '----������ ��������� �������
                MyTxt = "��� ������� ������ ��� ���������� ������������ ���� Excel, � ������� ������� �� ������ 11 �������: " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� A ������� ����� �� ������� (�� ������������ ��������), " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� B ������� ��� ������ Scala. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� C ������� ��� ������ ����������. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� D ������� �������� ������ . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� E ������� ������� ��������� ������ . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� F ���������� ���������� ������ . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� G ���������� ���� ������ ��� ���. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� I ���������� ���� �������� � �������. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "���� ����� � ������� - ���� �������� = 0. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ������� J ����� ���������� ��������� �������������. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "������ ������ ���� ��������� ��� ���������. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "��� ������� ������ ���� ���������, ����� B � C:" & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ��� ����� ������� ��� ��� ������ Scala, " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "��� ��� ������ ���������� (����� ����� ������� ���) " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "� ��� ���� �������������� ���� Excel � �� ������ ������ ������?" & Chr(13) & Chr(10)
                MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "��������!")
                If (MyRez = MsgBoxResult.Ok) Then
                    If My.Settings.UseOffice = "LibreOffice" Then
                        OpenFileDialog2.ShowDialog()
                        If (OpenFileDialog2.FileName = "") Then
                        Else
                            ImportFileName = OpenFileDialog2.FileName
                            Me.Cursor = Cursors.WaitCursor
                            Me.Refresh()
                            System.Windows.Forms.Application.DoEvents()
                            Declarations.MyOrderNum = Trim(Label3.Text)
                            ImportDataFromLO()
                        End If
                    Else
                        OpenFileDialog1.ShowDialog()
                        If (OpenFileDialog1.FileName = "") Then
                        Else
                            ImportFileName = OpenFileDialog1.FileName
                            Me.Cursor = Cursors.WaitCursor
                            Me.Refresh()
                            System.Windows.Forms.Application.DoEvents()
                            Declarations.MyOrderNum = Trim(Label3.Text)
                            ImportDataFromExcel()
                        End If
                    End If
                    Me.Cursor = Cursors.Default
                End If


            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ����� �������������� ���������� �� ���
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEDOInfo = New EDOInfo
        MyEDOInfo.ShowDialog()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ����� ����� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyShipmentsCost = New ShipmentsCost
        MyShipmentsCost.ShowDialog()
    End Sub
End Class