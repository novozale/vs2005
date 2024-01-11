Public Class CustomerSelectList

    Private Sub CustomerSelectList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список покупателей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(MyCustomerSelect.TextBox1.Text) = "" Then
            '----В первое окно условие не введено - считаем, что во второе введено
            'MySQLStr = "SELECT SL01001, SL01002, SL01003 + SL01004 + SL01005 AS SL01003, SL01021 "
            'MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
            MySQLStr = "SELECT SL010300.SL01001, SL010300.SL01002, SL010300.SL01003 + SL010300.SL01004 + SL010300.SL01005 AS SL01003, SL010300.SL01021, "
            MySQLStr = MySQLStr & "ISNULL(View_7.Address, N'') AS DelAddress "
            MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SL14001 AS Code, LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)) AS Address "
            MySQLStr = MySQLStr & "FROM SL140300 "
            MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_7 ON SL010300.SL01001 = View_7.Code "

            MySQLStr = MySQLStr & "WHERE (UPPER(SL01001) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(SL01002) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(SL01003 + SL01004 + SL01005) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(SL01021) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') "
            MySQLStr = MySQLStr & "ORDER BY SL01002"
        Else
            '----В первое окно условие введено
            If Trim(MyCustomerSelect.TextBox2.Text) = "" Then
                '----Во второе окно условие введено
                'MySQLStr = "SELECT SL01001, SL01002, SL01003 + SL01004 + SL01005 AS SL01003, SL01021 "
                'MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
                MySQLStr = "SELECT SL010300.SL01001, SL010300.SL01002, SL010300.SL01003 + SL010300.SL01004 + SL010300.SL01005 AS SL01003, SL010300.SL01021, "
                MySQLStr = MySQLStr & "ISNULL(View_7.Address, N'') AS DelAddress "
                MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SL14001 AS Code, LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)) AS Address "
                MySQLStr = MySQLStr & "FROM SL140300 "
                MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_7 ON SL010300.SL01001 = View_7.Code "

                MySQLStr = MySQLStr & "WHERE (UPPER(SL01001) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(SL01002) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(SL01003 + SL01004 + SL01005) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(SL01021) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & "ORDER BY SL01002"
            Else
                '----Условия введены в оба окна
                'MySQLStr = "SELECT SL01001, SL01002, SL01003 + SL01004 + SL01005 AS SL01003, SL01021 "
                'MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
                MySQLStr = "SELECT SL010300.SL01001, SL010300.SL01002, SL010300.SL01003 + SL010300.SL01004 + SL010300.SL01005 AS SL01003, SL010300.SL01021, "
                MySQLStr = MySQLStr & "ISNULL(View_7.Address, N'') AS DelAddress "
                MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SL14001 AS Code, LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)) AS Address "
                MySQLStr = MySQLStr & "FROM SL140300 "
                MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_7 ON SL010300.SL01001 = View_7.Code "

                MySQLStr = MySQLStr & "WHERE ((UPPER(SL01001) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(SL01001) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(SL01002) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(SL01002) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(SL01003 + SL01004 + SL01005) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(SL01003 + SL01004 + SL01005) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(SL01021) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(SL01021) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) "
                MySQLStr = MySQLStr & "ORDER BY SL01002 "
            End If

        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Код покупателя"
        DataGridView1.Columns(0).Width = 90
        DataGridView1.Columns(1).HeaderText = "Имя покупателя"
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(2).HeaderText = "Адрес покупателя"
        DataGridView1.Columns(2).Width = 400
        DataGridView1.Columns(3).HeaderText = "ИНН покупателя"
        DataGridView1.Columns(3).Width = 130
        DataGridView1.Columns(4).HeaderText = "Адрес доставки (умолч)"
        DataGridView1.Columns(4).Width = 400

        If DataGridView1.Rows.Count > 0 Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub CustomerSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MyCustomerSelect.DataGridView1.Rows.Count - 1
            If Trim(MyCustomerSelect.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MyCustomerSelect.DataGridView1.CurrentCell = MyCustomerSelect.DataGridView1.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub
End Class