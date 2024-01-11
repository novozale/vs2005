Public Class EditHeader

    Public StartParam As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с сохранением
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
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
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор валюты заказа
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
        '// Выбор валюты счета фактуры
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
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox2, True, True, True, False)
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox3, True, True, True, False)
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox4.Text = "Самовывоз со склада" Then
            Declarations.IsSelfDelivery = 1
        Else
            Declarations.IsSelfDelivery = 0
        End If
        Me.SelectNextControl(ComboBox4, True, True, True, False)
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub EditHeader_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
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
        Dim MyPRID As String                          'ID предложения
        Dim MyCCode As String                         'код покупателя
        Dim MyCName As String                         'имя покупателя
        Dim MyCAddr As String                         'адрес покупателя
        Dim MyWHNum As String                         'номер склада
        Dim MyDocCode As Integer                      'код документа
        Dim MyPRCurrCode As Integer                   'код валюты предложения
        Dim MyInvCurrCode As Integer                  'код валюты СФ
        Dim MyComment As String                       'примечание
        Dim MyPriceCond As String                     'Условия выставления цены
        Dim MyReadyDate As DateTime                   'дата готовности на складе
        Dim MyDeliveryAddr As String                  'адрес доставки
        Dim MyDeliveryDate As DateTime                'дата доставки
        Dim MyPaymentCond As String                   'условия платежа
        Dim MyExpirationDate As DateTime              'дата, ло которой действует предложение
        Dim MyPartialShipment As Integer              'Возможность частичной отгрузки (0 - нет, 1 - да)
        Dim MyAgentName As String                     'имя торгового агента, создавшего коммерческое предложение
        Dim MyCPState As String                       'состояние коммерческого предложения


        '----Добавление списка складов в ComboBox
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
            ComboBox1.DisplayMember = "SC23002" 'Это то что будет отображаться
            ComboBox1.ValueMember = "SC23001"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----Добавление списка кодов документов в ComboBox
        MySQLStr = "SELECT SY24002, SY24003 "
        MySQLStr = MySQLStr & "FROM SY240300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SY24001 = N'DC') "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT '' AS SY24002,'' AS SY24003 "
        MySQLStr = MySQLStr & "ORDER BY SY24002 "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.Fill(MyDs1)
            ComboBox6.DisplayMember = "SY24003" 'Это то что будет отображаться
            ComboBox6.ValueMember = "SY24002"   'это то что будет храниться
            ComboBox6.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----Добавление списка валют в ComboBox валюта документа
        MySQLStr = "SELECT SYCD001, SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD009 <> N'') "
        MySQLStr = MySQLStr & "AND (SYCD009 NOT IN ('FIM', 'FRF', 'SEK', 'DK', 'DM', 'FI1', 'ROL')) "
        Try
            MyAdapter3 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter3.Fill(MyDs3)
            ComboBox2.DisplayMember = "SYCD009" 'Это то что будет отображаться
            ComboBox2.ValueMember = "SYCD001"   'это то что будет храниться
            ComboBox2.DataSource = MyDs3.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----Добавление списка валют в ComboBox валюта СФ
        MySQLStr = "SELECT SYCD001, SYCD009 "
        MySQLStr = MySQLStr & "FROM SYCD0100 "
        MySQLStr = MySQLStr & "WHERE (SYCD009 <> N'') "
        MySQLStr = MySQLStr & "AND (SYCD009 = N'RUB') "
        MySQLStr = MySQLStr & "AND (SYCD009 NOT IN ('FIM', 'FRF', 'SEK', 'DK', 'DM', 'FI1', 'ROL')) "
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.Fill(MyDs4)
            ComboBox3.DisplayMember = "SYCD009" 'Это то что будет отображаться
            ComboBox3.ValueMember = "SYCD001"   'это то что будет храниться
            ComboBox3.DataSource = MyDs4.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----Добавление списка состояний КП в ComboBox
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_SalesmanWorkPlace4_CPState "
        MySQLStr = MySQLStr & "ORDER BY Name "
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.Fill(MyDs2)
            ComboBox7.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox7.ValueMember = "Name"   'это то что будет храниться
            ComboBox7.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        If StartParam = "Create" Then
            '----Получение номера предложения
            Label3.Text = Microsoft.VisualBasic.Right("0000000000" & CStr(GetNewID()), 10)
            ComboBox4.SelectedItem = "Самовывоз со склада"

        Else '----открытие на редактирование - заполнение существующих полей
            '----получение номера предложения
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
                '----Номер предложения (уже в MyPRID) 
                '----Код покупателя
                If Trim(Declarations.MyRec.Fields("SL01001").Value.ToString) = "" Then
                    MyCCode = Declarations.MyRec.Fields("OR01003").Value
                    TextBox2.ReadOnly = False
                    TextBox3.ReadOnly = False
                Else
                    MyCCode = Declarations.MyRec.Fields("SL01001").Value
                    TextBox2.ReadOnly = True
                    TextBox3.ReadOnly = True
                End If
                '----Имя покупателя
                If Declarations.MyRec.Fields("SL01002").Value.ToString = "" Then
                    MyCName = Declarations.MyRec.Fields("CName").Value
                Else
                    MyCName = Declarations.MyRec.Fields("SL01002").Value
                End If
                '----Адрес покупателя
                If Declarations.MyRec.Fields("SL01003").Value.ToString = "" Then
                    MyCAddr = Declarations.MyRec.Fields("CAddr").Value
                Else
                    MyCAddr = Declarations.MyRec.Fields("SL01003").Value
                End If
                '----Номер склада
                MyWHNum = Declarations.MyRec.Fields("OR01050").Value
                '----Код документа
                MyDocCode = Declarations.MyRec.Fields("DocCode").Value
                '----Валюта предложения
                MyPRCurrCode = Declarations.MyRec.Fields("OR01028").Value
                '----Валюта СФ
                MyInvCurrCode = Declarations.MyRec.Fields("OR01116").Value
                '----Примечание
                MyComment = Declarations.MyRec.Fields("OR17005").Value
                '----Условия выставления цены
                MyPriceCond = Declarations.MyRec.Fields("PriceCond").Value
                '----условия платежа
                MyPaymentCond = Declarations.MyRec.Fields("PaymentCond").Value
                '----дата, ло которой действует предложение
                MyExpirationDate = Declarations.MyRec.Fields("ExpirationDate").Value
                '----Возможность частичной отгрузки (0 - нет, 1 - да)
                MyPartialShipment = Declarations.MyRec.Fields("PartialShipment").Value
                '----имя торгового агента
                MyAgentName = Declarations.MyRec.Fields("AgentName").Value
                '----Состояние КП
                MyCPState = Declarations.MyRec.Fields("CPState").Value
                trycloseMyRec()

                '----Номер предложения
                Label3.Text = MyPRID
                '----Код покупателя
                TextBox1.Text = MyCCode
                '----Имя покупателя
                TextBox2.Text = MyCName
                '----Адрес покупателя
                TextBox3.Text = MyCAddr
                '----Номер склада
                ComboBox1.SelectedValue = MyWHNum
                '----Код документа
                ComboBox6.SelectedValue = MyDocCode
                '----Валюта предложения
                ComboBox2.SelectedValue = MyPRCurrCode
                '----Валюта СФ
                ComboBox3.SelectedValue = MyInvCurrCode
                '----Примечание
                TextBox4.Text = MyComment
                '----Условия выставления цены
                ComboBox4.Text = MyPriceCond
                If ComboBox4.Text = "Самовывоз со склада" Then
                    Declarations.IsSelfDelivery = 1
                Else
                    Declarations.IsSelfDelivery = 0
                End If
                '----условия платежа
                TextBox8.Text = MyPaymentCond
                '----дата, ло которой действует предложение
                DateTimePicker3.Value = MyExpirationDate
                '----Возможность частичной отгрузки (0 - нет, 1 - да)
                If MyPartialShipment = 0 Then
                    ComboBox5.Text = "Нет"
                Else
                    ComboBox5.Text = "Да"
                End If
                '----имя торгового агента
                TextBox5.Text = MyAgentName
                '----Состояние КП
                ComboBox7.SelectedValue = MyCPState
            End If
        End If

        '----Фокус в первое поле
        TextBox1.Select()
        Declarations.MyOrderNum = Trim(Label3.Text)
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна поиска покупателя
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
        '// Проверка - есть ли код в Scala, если да - то блокируем имя и адрес 
        '// Также проверяем - не заблокирован ли клиент в Scala
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
                MsgBox(MyRezStr, MsgBoxStyle.OkOnly, "Внимание!")
            End If
        Else
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
        End If
    End Sub

    Private Sub CheckCustomerBlock()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получаем информацию о блокировке клиента 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustomerIsCredit As Integer

        '---Получаем информацию по клиенту о блокировке
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
            If Declarations.MyRec.Fields("IsCredit").Value = 1 Then   '---кредитный клиент
                CustomerIsCredit = 1
            Else
                CustomerIsCredit = 0
            End If
            If Declarations.MyRec.Fields("IsBlocked").Value = 1 Then
                If Declarations.MyRec.Fields("IsCredit").Value = 1 Then   '---кредитный клиент
                    trycloseMyRec()
                    MsgBox("Клиент с кодом " & Trim(TextBox1.Text) & " является кредитным клиентом и в настоящий момент заблокирован. Основанием для блокировки является истечение срока кредитного договора и отсутствие в течении 2 - х лет платежей от данного клиента. Для такого клиента перевод заказа в 1 тип или выдача разрешения на отгрузку будет возможна только после занесения в карточку клиента информации по новому кредитному договору, новых реквизитов (при необходимости) и его разблокировка. " & _
                        "Для ввода новых реквизитов необходимо создать заявку 'Создание клиента' на портале. Для ввода информации по новому кредитному договору необходимо создать заявку 'Заключение кредитного договора' на портале. ", vbOKOnly, "Внимание!")
                Else                                                      '---некредитный клиент
                    trycloseMyRec()
                    MsgBox("Клиент с кодом " & Trim(TextBox1.Text) & " является некредитным клиентом и от него в течении 2 - х лет не было платежей. Для такого клиента перед созданием коммерческого предложения или заказа убедитесь, что реквизиты клиента не изменились. " & _
                        "Для ввода новых реквизитов в случае их изменения необходимо создать заявку 'Создание клиента' на портале. ", vbOKOnly, "Внимание!")
                End If
            Else
                If CustomerIsCredit = 1 Then                    '---проверяем только для кредитных покупателей
                    MySQLStr = "SELECT DataFrom, DataTo "
                    MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 "
                    MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(TextBox1.Text) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        MsgBox("Ошибка Salesman Workplace 4 функция CheckCustomerBlock Проверка - не вышел ли срок договора. Обратитесь к администратору.", vbCritical, "Внимание!")
                        trycloseMyRec()
                    Else
                        Declarations.MyRec.MoveFirst()
                        If Declarations.MyRec.Fields("DataFrom").Value = CDate("01/01/1900") Or Declarations.MyRec.Fields("DataTo").Value = CDate("01/01/1900") Then
                            '---даты не проставлены - не проверяем
                            trycloseMyRec()
                        Else
                            If Declarations.MyRec.Fields("DataFrom").Value <= Now() And Declarations.MyRec.Fields("DataTo").Value > Now() Then
                                '---Все OK, но проверим сколько осталость до окончания договора
                                If DateDiff("d", Now(), Declarations.MyRec.Fields("DataTo").Value) < 60 Then
                                    MsgBox("До конца действия договора с покупателем осталось меньше двух месяцев. Примите меры к заключению нового договора и занесите новые данные в базу.", vbOKOnly, "Внимание!")
                                End If
                                trycloseMyRec()
                            Else
                                MsgBox("Внимание! закончился или еще не начался срок действия текущего договора с данным клиентом. Заключите новый договор и занесите данные о нем в базу.", vbCritical, "Внимание!")
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
        '// Открытие окна редактирования срок заказа 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveHeader() = True Then
                '----открытие окна редактирования
                MyOrderLines = New OrderLines
                MyOrderLines.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckEmptyHDR() = True Then
            Me.Close()
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей формы
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Код покупателя"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If TextBox2.ReadOnly = False And Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Покупатель"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If TextBox3.ReadOnly = False And Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""Адрес покупателя"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = "" Then
            MsgBox("Склад отгрузки должен быть выбран", MsgBoxStyle.Critical, "Внимание!")
            ComboBox1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If ComboBox6.SelectedValue = "" Then
            MsgBox("Код документа должен быть выбран", MsgBoxStyle.Critical, "Внимание!")
            ComboBox6.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox2.Text) = "" Then
            MsgBox("Необходимо выбрать валюту предложения", MsgBoxStyle.Critical, "Внимание!")
            ComboBox2.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox3.Text) = "" Then
            MsgBox("Необходимо выбрать валюту счета - фактуры", MsgBoxStyle.Critical, "Внимание!")
            ComboBox3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox4.Text) = "" Then
            MsgBox("Необходимо выбрать условие выставления цен - введите стоимость доставки.", MsgBoxStyle.Critical, "Внимание!")
            Button7.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(TextBox8.Text) = "" Then
            MsgBox("Поле ""Условия оплаты"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            TextBox8.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If DateTimePicker3.Value < Now() Then
            MsgBox("Дата ""Срок действия предложения - до:"" должна быть больше текущей", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker3.Select()
            CheckFormFilling = False
            Exit Function
        End If

        If Trim(ComboBox5.Text) = "" Then
            MsgBox("Необходимо выбрать - возможна ли частичная поставка", MsgBoxStyle.Critical, "Внимание!")
            ComboBox5.Select()
            CheckFormFilling = False
            Exit Function
        End If

        CheckFormFilling = True
    End Function

    Private Function SaveHeader() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных введенных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyPRID As String                          'ID предложения
        Dim MyCCode As String                         'код покупателя
        Dim MyCName As String                         'имя покупателя
        Dim MyCAddr As String                         'адрес покупателя
        Dim MyWHNum As String                         'номер склада
        Dim MyDocCode As Integer                      'код документа
        Dim MyPRCurrCode As Integer                   'код валюты предложения
        Dim MyInvCurrCode As Integer                  'код валюты СФ
        Dim MyComment As String                       'примечание
        Dim MyPriceCond As String                     'Условия выставления цены
        Dim MyReadyDate As String                     'дата готовности на складе
        Dim MyDeliveryAddr As String                  'адрес доставки
        Dim MyDeliveryDate As String                  'дата доставки
        Dim MyPaymentCond As String                   'условия платежа
        Dim MyExpirationDate As String                'дата, ло которой действует предложение
        Dim MyPartialShipment As Integer              'Возможность частичной отгрузки (0 - нет, 1 - да)
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAgentName As String                     'имя торгового агента
        Dim MyCPState As String                       'Состояние КП

        '----Номер предложения
        MyPRID = Trim(Label3.Text)
        '----Код покупателя
        MyCCode = Trim(TextBox1.Text)
        '----Имя покупателя
        If TextBox2.Enabled = True Then
            MyCName = Trim(TextBox2.Text)
        Else
            MyCName = ""
        End If
        '----Адрес покупателя
        If TextBox3.Enabled = True Then
            MyCAddr = Trim(TextBox3.Text)
        Else
            MyCAddr = ""
        End If
        '----Номер склада
        MyWHNum = ComboBox1.SelectedValue
        '----код документа
        MyDocCode = ComboBox6.SelectedValue
        '----Валюта предложения
        MyPRCurrCode = Trim(ComboBox2.SelectedValue)
        '----Валюта СФ
        MyInvCurrCode = Trim(ComboBox3.SelectedValue)
        '----Примечание
        MyComment = Trim(TextBox4.Text)
        '----Условия выставления цены
        MyPriceCond = Trim(ComboBox4.Text)
        '----дата готовности на складе
        MyReadyDate = "01/01/1900"
        MyDeliveryAddr = ""
        '----дата доставки
        MyDeliveryDate = "01/01/1900"
        '----условия платежа
        MyPaymentCond = Trim(TextBox8.Text)
        '----дата, ло которой действует предложение
        MyExpirationDate = DatePart(DateInterval.Day, DateTimePicker3.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker3.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker3.Value)
        '----Возможность частичной отгрузки (0 - нет, 1 - да)
        If ComboBox5.Text = "Нет" Then
            MyPartialShipment = 0
        Else
            MyPartialShipment = 1
        End If
        '----имя торгового агента
        MyAgentName = TextBox5.Text
        '----Состояние КП
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
        '// Проверка - есть ли строки у предложения.
        '// Если нет - удаление заголовка и примечаний
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyPRID As String                          'ID предложения
        Dim MySQLStr As String                        'рабочая строка
        Dim MyRez As Double                           'результат - количество строк
        Dim MyRez1 As VariantType                     'результат выбора


        '----Номер предложения
        MyPRID = Trim(Label3.Text)
        MySQLStr = "SELECT COUNT(*) AS CL "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyPRID & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyRez = Declarations.MyRec.Fields("CL").Value
        trycloseMyRec()

        If MyRez = 0 Then '---нет строк в заказе
            MyRez1 = MsgBox("В предложении номер " & MyPRID & " нет ни одной строки. Удалить заголовок? ", MsgBoxStyle.YesNo, "Внимание!")
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
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox5, True, True, True, False)
    End Sub

    Private Sub DateTimePicker3_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.CloseUp
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker3, True, True, True, False)
    End Sub

    Private Sub DateTimePicker3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// импорт строк заказа из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String
        Dim MySQLStr As String

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveHeader() = True Then
                '----запуск процедуры импорта
                MyTxt = "Для импорта данных вам необходимо использовать файл Excel, в котором начиная со строки 11 указать: " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "в колонке A указать номер по порядку (не обязательный параметр), " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке B указать код запаса Scala. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке C указать код запаса поставщика. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке D указать название запаса . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке E выбрать единицу измерения запаса . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке F проставить количество запаса . " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке G проставить цену запаса без НДС. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке I проставить срок поставки в неделях. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "Если товар в наличии - срок поставки = 0. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "В колонке J можно проставить расчетную себестоимость. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "Строки должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "Все колонки должны быть заполнены, кроме B и C:" & Chr(13) & Chr(10)
                MyTxt = MyTxt & "в них можно указать или код товара Scala, " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "или код товара поставщика (также можно указать оба) " & Chr(13) & Chr(10)
                MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
                MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")
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
        '// открытие окна ввода дополнительной информации по ЭДО
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEDOInfo = New EDOInfo
        MyEDOInfo.ShowDialog()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна ввода суммы додставки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyShipmentsCost = New ShipmentsCost
        MyShipmentsCost.ShowDialog()
    End Sub
End Class