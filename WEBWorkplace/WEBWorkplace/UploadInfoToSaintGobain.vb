Public Class UploadInfoToSaintGobain

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Каталог для выгрузки обязательно должен быть выбран.", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub UploadInfoToSaintGobain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// запрет выхода по alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор каталога с картинками
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRange As Integer

        If Not TextBox1.Text.Equals("") Then
            Me.Cursor = Cursors.WaitCursor
            Select Case ComboBox1.SelectedItem
                Case "Все товары"
                    MyRange = 0
                Case "Все товары с картинками"
                    MyRange = 1
                Case "Все товары с картинками и описаниями"
                    MyRange = 2
                Case "Согласованный ассортимент"
                    MyRange = 3
                Case "Согласованный ассортимент с картинками"
                    MyRange = 4
                Case "Согласованный ассортимент с картинками и описаниями"
                    MyRange = 5
                Case Else
                    MyRange = 0
            End Select

            If My.Settings.UseOffice = "LibreOffice" Then
                UploadSenGobenToLO(TextBox1.Text, MyRange)
            Else
                UploadSenGobenToExcel(TextBox1.Text, MyRange)
            End If

            MsgBox("Выгрузка информации завершена", MsgBoxStyle.OkOnly, "Внимание!")
            Me.Cursor = Cursors.Default
        Else
            MsgBox("Необходимо выбрать каталог для выгрузки картинок", MsgBoxStyle.Critical, "Внимание!")
            Button1.Select()
        End If
    End Sub

    Private Sub UploadInfoToSaintGobain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ComboBox1.SelectedItem = "Согласованный ассортимент с картинками и описаниями"
    End Sub

End Class