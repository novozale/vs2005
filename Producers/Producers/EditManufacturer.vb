Public Class EditManufacturer
    Public NewManufacturer As Integer     'новый производитель (0) или редактирование (1)

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub EditManufacturer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в окно
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If NewManufacturer = 1 Then
            Declarations.MyManufacturerCode = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

            MySQLStr = "SELECT Name, Address, ContactInfo, Standard "
            MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyManufacturerCode & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("Данного производителя нет в БД. Возможно, он был удален кем то из пользователей. Обновите данные в окне производителей.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Me.Close()
            Else
                TextBox1.Text = Declarations.MyRec.Fields("Name").Value
                TextBox2.Text = Declarations.MyRec.Fields("Address").Value
                TextBox3.Text = Declarations.MyRec.Fields("ContactInfo").Value
                CheckBox1.Checked = Declarations.MyRec.Fields("Standard").Value
                trycloseMyRec()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckData() = True Then
            If NewManufacturer = 1 Then '------------редактирование записи
                MySQLStr = "UPDATE tbl_Manufacturers "
                MySQLStr = MySQLStr & "SET Name = N'" & GetSQLStrng(Trim(TextBox1.Text)) & "', "
                MySQLStr = MySQLStr & "Address = N'" & GetSQLStrng(Trim(TextBox2.Text)) & "', "
                MySQLStr = MySQLStr & "ContactInfo = N'" & GetSQLStrng(Trim(TextBox3.Text)) & "', "
                If CheckBox1.Checked = True Then
                    MySQLStr = MySQLStr & "Standard = -1 "
                Else
                    MySQLStr = MySQLStr & "Standard = 0 "
                End If
                MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyManufacturerCode & ")"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else                        '------------создание записи
                MySQLStr = "INSERT INTO tbl_Manufacturers "
                MySQLStr = MySQLStr & "(Name, Address, ContactInfo, Standard) "
                MySQLStr = MySQLStr & "VALUES (N'" & GetSQLStrng(Trim(TextBox1.Text)) & "', "
                MySQLStr = MySQLStr & "N'" & GetSQLStrng(Trim(TextBox2.Text)) & "', "
                MySQLStr = MySQLStr & "N'" & GetSQLStrng(Trim(TextBox3.Text)) & "', "
                If CheckBox1.Checked = True Then
                    MySQLStr = MySQLStr & "-1 ) "
                Else
                    MySQLStr = MySQLStr & "0 ) "
                End If
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If
            Me.Close()
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка ввода данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Название"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckData = False
            Exit Function
        End If

        'If Trim(TextBox2.Text) = "" Then
        '    MsgBox("Поле ""Адрес"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
        '    TextBox2.Select()
        '    CheckData = False
        '    Exit Function
        'End If
        CheckData = True
    End Function
End Class