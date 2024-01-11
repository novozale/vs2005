Public Class CommentAndCancelReason
    Public MyID As Integer


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна и запись выбранного
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr + "Set Comments = ISNULL(Comments, '') + " + Chr(10) + Chr(13) + " + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & Trim(TextBox1.Text) & "', "
        MySQLStr = MySQLStr + "CancelReason = N'" & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString) & "' "
        MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Me.Close()
    End Sub

    Private Sub CommentAndCancelReason_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub CommentAndCancelReason_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка причин отказов
        Dim MyDs As New DataSet

        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_CancelReasons "
        MySQLStr = MySQLStr & "ORDER BY Name"

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Причина"
    End Sub
End Class