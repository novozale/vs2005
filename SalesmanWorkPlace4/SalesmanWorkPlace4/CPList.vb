Public Class CPList

    Private Sub CPList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна и данных в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////
        
        DateTimePicker1.Value = DateAdd(DateInterval.Quarter, -2, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
        DateTimePicker2.Value = CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))

        '---Вывод данных в окно
        CPDataLoad()
        CheckButtons()
    End Sub

    Private Sub CPDataLoad()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных по КП
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet

        MySQLStr = "SELECT View_1.OR01001, View_1.OR01015, View_1.ExpirationDate, View_1.OR01003, CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') "
        MySQLStr = MySQLStr & "+ ' ' + ISNULL(SL010300.SL01003, N''))) = '' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, "
        MySQLStr = MySQLStr & "N'') END AS CName, ISNULL(View_1.AgentName, N'') AS AgentName, View_1.OrderN, ISNULL(View_3.Comment, N'') AS Comment "
        MySQLStr = MySQLStr & "FROM (SELECT OR17001, LTRIM(RTRIM(LTRIM(RTRIM(OR17005)) + ' ' + LTRIM(RTRIM(OR17006)))) AS Comment "
        MySQLStr = MySQLStr & "FROM  tbl_OR170300) AS View_3 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR01001, OR01003, OR01015, OrderN, CName, ExpirationDate, AgentName "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR01096 = N'" & Declarations.SalesmanCode & "') "
        MySQLStr = MySQLStr & "AND (OR01015 >= CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)) "
        MySQLStr = MySQLStr & "AND (OR01015 <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', "
        MySQLStr = MySQLStr & "103))) AS View_1 ON View_3.OR17001 = View_1.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 WITH (NOLOCK) ON View_1.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "ORDER BY View_1.OR01001 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "Номер предложения"
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).HeaderText = "Дата создания"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).HeaderText = "Действительно до"
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).HeaderText = "Код покупателя"
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).HeaderText = "Имя покупателя"
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).HeaderText = "Имя агента"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(5).ReadOnly = True
        If My.Settings.ShowAgentColumn = 0 Then
            DataGridView1.Columns(5).Visible = False
        Else
            DataGridView1.Columns(5).Visible = True
        End If
        DataGridView1.Columns(6).HeaderText = "Перенесено в заказ на продажу номер"
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).HeaderText = "Комментарий"
        DataGridView1.Columns(7).Width = 200
        DataGridView1.Columns(7).ReadOnly = True
    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button1.Enabled = False
 
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При изменении даты "с" и "по" перезагружаем данные
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CPDataLoad()
        CheckButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора КП
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором КП
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyCPID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        Me.Close()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором КП
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyCPID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        Me.Close()
    End Sub
End Class