Public Class Salesman

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Salesman_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet                       '

        '---------------Список городов
        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Cities "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT 0 AS ID, '' AS Name "
        MySQLStr = MySQLStr & "ORDER BY ID "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "ID"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '--------------Собственно данные по продавцу
        MySQLStr = "SELECT Code, Name, Email, ISNULL(City,0) AS City, OfficeLeader, OnDuty, IsActive, Rezerv1, Rezerv2 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Salesmans WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & UCase(Trim(Declarations.MySalesmanID)) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Выделенный продавец не найден, возможно удален другим пользователем. закройте и откройте по новой вкладку продавцов.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Me.Close()
        Else
            TextBox1.Text = Declarations.MyRec.Fields("Code").Value.ToString
            TextBox2.Text = Declarations.MyRec.Fields("Name").Value
            TextBox3.Text = Declarations.MyRec.Fields("Email").Value
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("City").Value
            If Declarations.MyRec.Fields("OfficeLeader").Value = 0 Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
            If Declarations.MyRec.Fields("OnDuty").Value = 0 Then
                CheckBox2.Checked = False
            Else
                CheckBox2.Checked = True
            End If
            If Declarations.MyRec.Fields("IsActive").Value = 0 Then
                CheckBox3.Checked = False
            Else
                CheckBox3.Checked = True
            End If
            TextBox4.Text = Declarations.MyRec.Fields("Rezerv1").Value
            TextBox5.Text = Declarations.MyRec.Fields("Rezerv2").Value
            trycloseMyRec()
        End If
        TextBox1.Enabled = False
        TextBox2.Enabled = False
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckBox3.Checked = True Then
            If Trim(TextBox3.Text) = "" Then
                MsgBox("Поле ""Электронная почта"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
                TextBox3.Select()
                CheckData = False
                Exit Function
            End If

            If ComboBox1.SelectedValue = 0 Then
                MsgBox("""Город продавца"" должен быть выбран", MsgBoxStyle.Critical, "Внимание!")
                ComboBox1.Select()
                CheckData = False
                Exit Function
            End If
        End If

        CheckData = True
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckData() = True Then
            MySQLStr = "UPDATE tbl_WEB_Salesmans "
            MySQLStr = MySQLStr & "SET Email = N'" & Trim(TextBox3.Text) & "', "
            MySQLStr = MySQLStr & "City = " & ComboBox1.SelectedValue & ", "
            If CheckBox1.Checked = True Then
                MySQLStr = MySQLStr & "OfficeLeader = N'1', "
            Else
                MySQLStr = MySQLStr & "OfficeLeader = N'0', "
            End If
            If CheckBox2.Checked = True Then
                MySQLStr = MySQLStr & "OnDuty = N'1', "
            Else
                MySQLStr = MySQLStr & "OnDuty = N'0', "
            End If
            If CheckBox3.Checked = True Then
                MySQLStr = MySQLStr & "IsActive = 1, "
            Else
                MySQLStr = MySQLStr & "IsActive = 0, "
            End If
            MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(TextBox4.Text) & "', "
            MySQLStr = MySQLStr & "Rezerv2 = N'" & Trim(TextBox5.Text) & "' "
            If CheckBox3.Checked = True Then
                MySQLStr = MySQLStr & ", RMStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE CASE WHEN ScalaStatus = 2 THEN 2 ELSE 3 END END "
                MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN ScalaStatus = 1 THEN 1 ELSE CASE WHEN ScalaStatus = 2 THEN 2 ELSE 3 END END "
            Else
                MySQLStr = MySQLStr & ", RMStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE CASE WHEN IsActive = 1 THEN 2 ELSE 3 END END "
                MySQLStr = MySQLStr & ", WEBStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE CASE WHEN IsActive = 1 THEN 2 ELSE 3 END END "
            End If
            MySQLStr = MySQLStr & "WHERE (Code = N'" & CStr(Declarations.MySalesmanID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            Me.Close()
        End If
    End Sub
End Class