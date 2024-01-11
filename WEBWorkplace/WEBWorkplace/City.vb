Public Class City
    Public StartParam As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckData() = True Then
            If StartParam = "Edit" Then '---обновление данных
                MySQLStr = "UPDATE tbl_WEB_Cities "
                MySQLStr = MySQLStr & "SET Name = N'" & Trim(TextBox2.Text) & "' "
                MySQLStr = MySQLStr & "WHERE (ID = " & CStr(Declarations.MyCityID) & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else                        '---запись новых данных
                MySQLStr = "INSERT INTO tbl_WEB_Cities "
                MySQLStr = MySQLStr & "(ID, Name) "
                MySQLStr = MySQLStr & "VALUES (" & CStr(Declarations.MyCityID) & ", "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox2.Text) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub City_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If StartParam = "Edit" Then
            MySQLStr = "SELECT ID, Name "
            MySQLStr = MySQLStr & "FROM  tbl_WEB_Cities WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ID = " & UCase(Trim(Declarations.MyCityID)) & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Выделенный город не найден, возможно удален другим пользователем. закройте и откройте по новой вкладку городов.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Me.Close()
            Else
                TextBox1.Text = Declarations.MyRec.Fields("ID").Value.ToString
                TextBox2.Text = Declarations.MyRec.Fields("Name").Value
                trycloseMyRec()
            End If
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение  кода города в глобальной переменной
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) <> "" Then
            Declarations.MyCityID = CInt(TextBox1.Text)
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка ввода кода города
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Integer
        Dim MySQLStr As String                        'рабочая строка

        If Trim(TextBox1.Text) <> "" Then
            Try
                MyRez = TextBox1.Text
            Catch ex As Exception
                MsgBox("В поле ""Код города"" должно быть введено целое положительное число больше 0.", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            End Try

            If MyRez <= 0 Then
                MsgBox("В поле ""Код города"" должно быть введено целое положительное число больше 0.", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            End If

            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Cities "
            MySQLStr = MySQLStr & "WHERE (ID = " & TextBox1.Text & ") "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                e.Cancel = False
            Else
                If Declarations.MyRec.Fields("CC").Value > 0 Then
                    MsgBox("Код города " & TextBox1.Text & " уже есть в базе данных. Введите другой код.", MsgBoxStyle.Critical, "Внимание!")
                    trycloseMyRec()
                    e.Cancel = True
                Else
                    trycloseMyRec()
                    e.Cancel = False
                End If
            End If
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Код города"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Название города"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            Exit Function
        End If

        CheckData = True
    End Function
End Class