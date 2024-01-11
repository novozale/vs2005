Public Class CustomerSelect
    Public StartParam As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна без выбора покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CustomerSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список покупателей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        'MySQLStr = "SELECT SL01001, SL01002, SL01003 + SL01004 + SL01005 AS SL01003, SL01021 "
        'MySQLStr = MySQLStr & "FROM SL01" & Declarations.CompanyID & "00 WITH(NOLOCK) "
        'MySQLStr = MySQLStr & "ORDER BY SL01002 "
        MySQLStr = "SELECT SL010300.SL01001, SL010300.SL01002, SL010300.SL01003 + SL010300.SL01004 + SL010300.SL01005 AS SL01003, SL010300.SL01021, "
        MySQLStr = MySQLStr & "ISNULL(View_7.Address, N'') AS DelAddress "
        MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SL14001 AS Code, LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)) AS Address "
        MySQLStr = MySQLStr & "FROM SL140300 "
        MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_7 ON SL010300.SL01001 = View_7.Code "
        MySQLStr = MySQLStr & "ORDER BY SL010300.SL01002 "

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

        CheckButtons()
    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "Подсветить все"
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка состояния кнопок при изменении выделения  
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub CustomerSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As Double
        Dim MyRezStr As String

        If StartParam = "CP" Then    '-----коммерческое предложение
            MyEditHeader.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(MyEditHeader.TextBox1.Text) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            MyRez = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
            If MyRez = 1 Then
                MyRezStr = CheckSalesman(Declarations.SalesmanCode, Trim(MyEditHeader.TextBox1.Text))
                If MyRezStr = "" Then
                    MyEditHeader.TextBox2.ReadOnly = True
                    MyEditHeader.TextBox3.ReadOnly = True
                    MySQLStr = "SELECT SL01002, SL01003 + ' ' + SL01004 + ' ' + SL01005 AS SL01003, "
                    MySQLStr = MySQLStr & "SL01085, SL01098 "
                    MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(MyEditHeader.TextBox1.Text) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    MyEditHeader.TextBox2.Text = Declarations.MyRec.Fields("SL01002").Value
                    MyEditHeader.TextBox3.Text = Declarations.MyRec.Fields("SL01003").Value
                    MyEditHeader.ComboBox6.SelectedValue = Declarations.MyRec.Fields("SL01085").Value
                    MyEditHeader.ComboBox1.SelectedValue = Declarations.MyRec.Fields("SL01098").Value
                    trycloseMyRec()
                Else
                    MyEditHeader.TextBox1.Text = ""
                    MyEditHeader.TextBox2.Text = ""
                    MyEditHeader.TextBox3.Text = ""
                    MyEditHeader.TextBox2.ReadOnly = False
                    MyEditHeader.TextBox3.ReadOnly = False
                    MsgBox(MyRezStr, MsgBoxStyle.OkOnly, "Внимание!")
                End If
            Else
                MyEditHeader.TextBox2.ReadOnly = False
                MyEditHeader.TextBox3.ReadOnly = False
            End If
        ElseIf StartParam = "Search" Then    '-----создание запроса на поиск поставщика
            MyEditRequest.TextBox1.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            MyEditRequest.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
            MySQLStr = "SELECT SL230300.SL23004 AS PT "
            MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
            MySQLStr = MySQLStr & "SL230300 ON SL010300.SL01024 = SL230300.SL23003 "
            MySQLStr = MySQLStr & "WHERE (SL230300.SL23002 = N'RUS') "
            MySQLStr = MySQLStr & "AND (SL230300.SL23001 = N'0') "
            MySQLStr = MySQLStr & "AND (SL010300.SL01001 = N'" & Trim(MyEditRequest.TextBox1.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyEditRequest.TextBox9.Text = ""
                MyEditRequest.TextBox9.Enabled = True
            Else
                MyEditRequest.TextBox9.Text = Declarations.MyRec.Fields("PT").Value.ToString
                MyEditRequest.TextBox9.Enabled = False
            End If
            trycloseMyRec()
        End If
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего по критерию покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего по критерию покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit Sub
                    End If
                End If
            Next i
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсвечивание всех подходящих по критерию покупателей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            Button6.Text = "Снять выделение"
        Else
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            Button6.Text = "Подсветить все"
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию покупателей в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox1.Select()
        Else
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub
End Class