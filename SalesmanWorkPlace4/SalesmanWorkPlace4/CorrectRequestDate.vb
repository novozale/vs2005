Public Class CorrectRequestDate
    Public MyID As Integer
    Public MyDate As DateTime

    Private Sub CorrectRequestDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub CorrectRequestDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DateTimePicker1.Value = MyDate
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с сохранением
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----сохранение результатов
            If SaveRequest() = True Then
                Me.Close()
            End If
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей формы
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----дата больше или равна текущей
        If DateTimePicker1.Value < Now().AddDays(-1) Then
            MsgBox("Дата ""Срок представления КП"" должна быть больше или равна текущей", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        '-----после 15:00 дата только больше текущей
        If Hour(Now()) >= 15 Then
            If DateTimePicker1.Value < Now() Then
                MsgBox("После 15:00 Дата ""Срок представления КП"" должна быть больше текущей", MsgBoxStyle.Critical, "Внимание!")
                DateTimePicker1.Select()
                CheckFormFilling = False
                Exit Function
            End If
        End If

        '-----нельзя выбирать субботу и воскресенье
        If Weekday(DateTimePicker1.Value, 2) = 6 Or Weekday(DateTimePicker1.Value, 2) = 7 Then
            MsgBox("Дата ""Срок представления КП"" не должна быть субботой или воскресеньем.", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        CheckFormFilling = True
    End Function

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных введенных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr & "SET RequestDate = CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103) "
        MySQLStr = MySQLStr & "WHERE (ID = " & CStr(MyID) & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        SaveRequest = True
    End Function
End Class