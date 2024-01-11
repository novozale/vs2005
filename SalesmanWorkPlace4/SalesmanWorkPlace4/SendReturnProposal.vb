Public Class SendReturnProposal

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из формы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SendReturnProposal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы, загрузка данных
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    '
        Dim MyDs As New DataSet                       '
        Dim MyDs1 As New DataSet                      '
        Dim MyDs2 As New DataSet                      '

        ''---Определение есть ли еще продавцы в группе
        'MySQLStr = "SELECT COUNT(tbl_SalesCommission_Groups.SalesmanCode) AS CC "
        'MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName "
        'MySQLStr = MySQLStr & "WHERE (tbl_SalesCommission_Groups_1.SalesmanCode = N'" & Declarations.SalesmanCode & "') "
        'MySQLStr = MySQLStr & "AND (tbl_SalesCommission_Groups.SalesmanCode <> N'" & Declarations.SalesmanCode & "') "
        'InitMyConn(False)
        'InitMyRec(False, MySQLStr)
        'Declarations.MyRec.MoveFirst()
        'If Declarations.MyRec.Fields("CC").Value = 0 Then
        '    trycloseMyRec()
        '    MsgBox("передача заказов возможна только внутри группы продавцов. Для вас группа продавцов не поределена. Обратитесь к администратору.", vbCritical, "Внимание!")
        '    Me.Close()
        'Else
        '    trycloseMyRec()
        'End If

        '=============Список продавцов в комбобоксы==========================================
        '-------------От кого передать-------------------------------------------------------
        'MySQLStr = "SELECT ST010300.ST01001 AS Code, Ltrim(Rtrim(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        'MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName INNER JOIN "
        'MySQLStr = MySQLStr & "ST010300 ON tbl_SalesCommission_Groups_1.SalesmanCode = ST010300.ST01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_SalesCommission_Groups.SalesmanCode = N'" & Declarations.SalesmanCode & "') "
        'MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        MySQLStr = "SELECT ST010300.ST01001 AS Code, Ltrim(Rtrim(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "Code"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox1.SelectedValue = Declarations.SalesmanCode

        '-------------Кому передать-------------------------------------------------------
        'MySQLStr = "SELECT ST010300.ST01001 AS Code, LTRIM(RTRIM(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        'MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName INNER JOIN "
        'MySQLStr = MySQLStr & "ST010300 ON tbl_SalesCommission_Groups_1.SalesmanCode = ST010300.ST01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_SalesCommission_Groups.SalesmanCode = N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "AND (tbl_SalesCommission_Groups_1.SalesmanCode <> N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        MySQLStr = "SELECT ST010300.ST01001 AS Code, LTRIM(RTRIM(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.SelectCommand.CommandTimeout = 600
            MyAdapter2.Fill(MyDs2)
            ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox2.ValueMember = "Code"   'это то что будет храниться
            ComboBox2.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '-------------Кому вернуть-------------------------------------------------------
        'MySQLStr = "SELECT ST010300.ST01001 AS Code, Ltrim(Rtrim(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        'MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName INNER JOIN "
        'MySQLStr = MySQLStr & "ST010300 ON tbl_SalesCommission_Groups_1.SalesmanCode = ST010300.ST01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_SalesCommission_Groups.SalesmanCode = N'" & Declarations.SalesmanCode & "') "
        'MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        MySQLStr = "SELECT ST010300.ST01001 AS Code, Ltrim(Rtrim(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox5.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox5.ValueMember = "Code"   'это то что будет храниться
            ComboBox5.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox5.SelectedValue = Declarations.SalesmanCode

        GetTransferredOrders()
        GetTransferredBackOrders()
    End Sub

    Private Sub GetTransferredOrders()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение списка передаваемых заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ' Все' AS OrderN, 'Все заказы' AS OrderInfo "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT tbl_OR010300.OR01001 AS OrderN, Ltrim(Rtrim(tbl_OR010300.OR01001)) + ' ' + CONVERT(nvarchar(30), tbl_OR010300.OR01015, 103) + ' ' + LTRIM(RTRIM(tbl_OR010300.OR01003)) "
        MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(SL010300.SL01002, ''))) AS OrderInfo "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON tbl_OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01019 = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (tbl_OR010300.OR01096 = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "ORDER BY OrderN"
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox3.DisplayMember = "OrderInfo" 'Это то что будет отображаться
            ComboBox3.ValueMember = "OrderN"   'это то что будет храниться
            ComboBox3.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox3.SelectedValue = " Все"
    End Sub

    Private Sub GetTransferredBackOrders()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение списка возвращаемых заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ' Все' AS OrderN, 'Все заказы' AS OrderInfo "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT tbl_OR010300.OR01001 AS OrderN, Ltrim(Rtrim(tbl_OR010300.OR01001)) + ' ' + CONVERT(nvarchar(30), tbl_OR010300.OR01015, 103) + ' ' + LTRIM(RTRIM(tbl_OR010300.OR01003)) "
        MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(SL010300.SL01002, ''))) AS OrderInfo "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON tbl_OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01019 = N'" & ComboBox5.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (tbl_OR010300.OR01096 <> N'" & ComboBox5.SelectedValue & "') "
        MySQLStr = MySQLStr & "ORDER BY OrderN"
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox4.DisplayMember = "OrderInfo" 'Это то что будет отображаться
            ComboBox4.ValueMember = "OrderN"   'это то что будет храниться
            ComboBox4.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox4.SelectedValue = " Все"
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор продавца, от кого передаются заказы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button1.Select()
    End Sub

    Private Sub ComboBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ComboBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение списка продавцов, кому передаются заказы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---список, кому передавать
        'MySQLStr = "SELECT ST010300.ST01001 AS Code, LTRIM(RTRIM(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        'MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
        'MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName INNER JOIN "
        'MySQLStr = MySQLStr & "ST010300 ON tbl_SalesCommission_Groups_1.SalesmanCode = ST010300.ST01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_SalesCommission_Groups.SalesmanCode = N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "AND (tbl_SalesCommission_Groups_1.SalesmanCode <> N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        MySQLStr = "SELECT ST010300.ST01001 AS Code, LTRIM(RTRIM(ST010300.ST01001)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS Name "
        MySQLStr = MySQLStr & "FROM ST010300 "
        MySQLStr = MySQLStr & "ORDER BY ST010300.ST01002 "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox2.ValueMember = "Code"   'это то что будет храниться
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---список передаваемых заказов
        GetTransferredOrders()
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор продавца, кому передаются заказы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button1.Select()
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор продавца, от которого возвращаются заказы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button2.Select()
    End Sub

    Private Sub ComboBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ComboBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбран продавец, от которого возвращаются заказы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        GetTransferredBackOrders()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передача заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(ComboBox3.SelectedValue) = "Все" Then
            MySQLStr = "UPDATE tbl_OR010300 "
            MySQLStr = MySQLStr & "Set OR01096 = N'" & Trim(ComboBox2.SelectedValue) & "' "
            MySQLStr = MySQLStr & "WHERE (OR01019 = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "AND (OR01096 = N'" & Trim(ComboBox1.SelectedValue) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MsgBox("Передача владения заказами завершена", vbOKOnly, "Внимание!")
        Else
            MySQLStr = "UPDATE tbl_OR010300 "
            MySQLStr = MySQLStr & "Set OR01096 = N'" & Trim(ComboBox2.SelectedValue) & "' "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(ComboBox3.SelectedValue) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MsgBox("Передача владения заказами завершена", vbOKOnly, "Внимание!")
        End If
        GetTransferredOrders()
        GetTransferredBackOrders()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Возврат заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(ComboBox4.SelectedValue) = "Все" Then
            MySQLStr = "UPDATE tbl_OR010300 "
            MySQLStr = MySQLStr & "Set OR01096 = OR01019 "
            MySQLStr = MySQLStr & "WHERE (OR01096 <> OR01019) "
            MySQLStr = MySQLStr & "AND (OR01019 = N'" & Trim(ComboBox5.SelectedValue) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MsgBox("Передача владения заказами завершена", vbOKOnly, "Внимание!")
        Else
            MySQLStr = "UPDATE tbl_OR010300 "
            MySQLStr = MySQLStr & "Set OR01096 = OR01019 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(ComboBox4.SelectedValue) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MsgBox("Передача владения заказами завершена", vbOKOnly, "Внимание!")
        End If
        GetTransferredOrders()
        GetTransferredBackOrders()
    End Sub
End Class