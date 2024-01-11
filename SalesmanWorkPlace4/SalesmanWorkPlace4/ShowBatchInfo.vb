Public Class ShowBatchInfo
    Public MyItem As String                           'код запаса

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ShowBatchInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "SELECT SC01001 + ' ' + SC01002 as Name "
        MySQLStr = MySQLStr & "FROM SC010300 "
        MySQLStr = MySQLStr & "where SC01001 = N'" & Trim(MyItem) & "'"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        Label2.Text = Declarations.MyRec.Fields("Name").Value
        trycloseMyRec()

        MySQLStr = "SELECT LTRIM(RTRIM(SC330300.SC33002)) + ' ' + LTRIM(RTRIM(SC230300.SC23002)) AS WH, "
        MySQLStr = MySQLStr & "SC330300.SC33003 AS BatchNum, "
        MySQLStr = MySQLStr & "SC330300.SC33004 AS BinNum, "
        MySQLStr = MySQLStr & "SC330300.SC33005 AS Balance, "
        MySQLStr = MySQLStr & "SC330300.SC33005 - SC330300.SC33006 - SC330300.SC33007 AS Available, "
        MySQLStr = MySQLStr & "SC330300.SC33006 AS Allocated, "
        MySQLStr = MySQLStr & "SC330300.SC33007 AS Ordered "
        MySQLStr = MySQLStr & "FROM SC330300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC230300 ON SC330300.SC33002 = SC230300.SC23001 "
        MySQLStr = MySQLStr & "WHERE (SC330300.SC33001 = N'" & Trim(MyItem) & "') AND "
        MySQLStr = MySQLStr & "(SC330300.SC33005 <> 0) AND "
        MySQLStr = MySQLStr & "(SC230300.SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY WH, BatchNum "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "Склад"
        DataGridView1.Columns(1).HeaderText = "N партии"
        DataGridView1.Columns(2).HeaderText = "Ячейка"
        DataGridView1.Columns(3).HeaderText = "Баланс"
        DataGridView1.Columns(4).HeaderText = "Доступно"
        DataGridView1.Columns(5).HeaderText = "Распределено"
        DataGridView1.Columns(6).HeaderText = "Задолж."
    End Sub
End Class