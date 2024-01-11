Public Class BasePrice

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранени€
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub BasePrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабоча€ строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'дл€ списка подгрупп
        Dim MyDs As New DataSet

        '---------------—писок базовых прайс листов
        MySQLStr = "SELECT DISTINCT SY240300.SY24002 AS Code, SY240300.SY24002 + ' ' + SY240300.SY24003 AS Name "
        MySQLStr = MySQLStr & "FROM SY240300 INNER JOIN "
        MySQLStr = MySQLStr & "SC390300 ON SY240300.SY24002 = SC390300.SC39002 "
        MySQLStr = MySQLStr & "WHERE (SY240300.SY24001 = N'IL') "
        MySQLStr = MySQLStr & "ORDER BY SY240300.SY24002 "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Ёто то что будет отображатьс€
            ComboBox1.ValueMember = "Code"   'это то что будет хранитьс€
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ¬ыгрузка в Excel базового прайс листа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadBasePriceToLO(ComboBox1.SelectedValue, ComboBox1.Text, CheckBox1.Checked)
        Else
            UploadBasePriceToExcel(ComboBox1.SelectedValue, ComboBox1.Text, CheckBox1.Checked)
        End If
        Me.Cursor = Cursors.Default
    End Sub
End Class