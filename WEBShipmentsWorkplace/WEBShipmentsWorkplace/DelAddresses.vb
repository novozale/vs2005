Public Class DelAddresses

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub DelAddresses_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без выбора по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub DelAddresses_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список адресов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        LoadData()
        CheckButtons()
    End Sub

    Private Sub LoadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка адресов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT     SL14002 AS ID, LTRIM(RTRIM(LTRIM(RTRIM(SL14003)) + ' ' + LTRIM(RTRIM(SL14004)) + ' ' + LTRIM(RTRIM(SL14005)) + ' ' + LTRIM(RTRIM(SL14006)))) AS Address "
        MySQLStr = MySQLStr & "FROM SL140300 "
        MySQLStr = MySQLStr & "WHERE (SL14001 = N'" & Trim(MyShipment.LblCustomerCode.Text) & "') "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(1).HeaderText = "Адрес"
        DataGridView1.Columns(1).Width = 700
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление списка адресов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = Trim(MyShipment.LblCustomerCode.Text) + " " + Trim(MyShipment.LblCustomerName.Text)
        LoadData()
        CheckButtons()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с заменой адреса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            AddressSelect()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором адреса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        AddressSelect()
    End Sub

    Private Sub AddressSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором адреса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyShipment.TextBox2.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        Me.Close()
    End Sub
End Class