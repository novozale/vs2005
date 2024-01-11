Public Class EstimatedIncome
    Public MyItem As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна с информацией об ожидаемом приходе
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub EstimatedIncome_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        Label2.Text = MyShowWHDetails.Label2.Text
        MySQLStr = "SELECT PC030300.PC03001, PC010300.PC01023, PC010300.PC01060 AS SalesOrderNum, PC030300.PC03010 - PC030300.PC03011 AS QTY, "
        MySQLStr = MySQLStr & "PC030300.PC03024, PC030300.PC03031 "
        MySQLStr = MySQLStr & "FROM PC030300 INNER JOIN "
        MySQLStr = MySQLStr & "PC010300 ON PC030300.PC03001 = PC010300.PC01001 "
        MySQLStr = MySQLStr & "WHERE (PC030300.PC03005 = N'" & Trim(MyItem) & "') "
        MySQLStr = MySQLStr & "AND (PC030300.PC03010 - PC030300.PC03011 > 0) "
        MySQLStr = MySQLStr & "ORDER BY PC03024 "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "N заказа на закупку"
        DataGridView1.Columns(0).Width = 150
        DataGridView1.Columns(1).HeaderText = "Склад"
        DataGridView1.Columns(2).HeaderText = "N заказа на продажу"
        DataGridView1.Columns(3).HeaderText = "Количество"
        DataGridView1.Columns(4).HeaderText = "Ожидаемая дата поставки"
        DataGridView1.Columns(4).DefaultCellStyle.Format = "dd/MM/yyyy"
        DataGridView1.Columns(5).HeaderText = "Задолженная дата поставки"
        DataGridView1.Columns(5).DefaultCellStyle.Format = "dd/MM/yyyy"
    End Sub
End Class