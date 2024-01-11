Public Class UploadPicturesToMagento

    Private Sub UploadPicturesToMagento_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub UploadPicturesToMagento_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    'для поставщиков
        Dim MyDs1 As New DataSet

        InitMyConn(False)
        '---поставщики
        MySQLStr = "SELECT '---' AS Code, 'Все' AS Name "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT DISTINCT SC010300.SC01058 AS Code, SC010300.SC01058 + ' ' + PL010300.PL01002 AS Name "
        MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "ORDER BY Code "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "Code"   'это то что будет храниться
            ComboBox1.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без загрузки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации на сайт Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        Button2.Enabled = False
        Button3.Enabled = False
        GroupBox1.BackColor = Color.LightGray
        Label2.Text = "0"
        Label3.Text = "0"
        UploadPictures_ToMagento(ComboBox1.SelectedValue)
        Me.Cursor = Cursors.Default
        Button2.Enabled = True
        Button3.Enabled = True
        MyErrWindow = New ErrWindow
        MyErrWindow.ShowDialog()
    End Sub
End Class