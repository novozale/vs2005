Public Class ErrForm
    Public MyErrStr As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с результатом 0 (прекращение дальнейшей работы)  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyErrRezult = 0
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с результатом 1 (продолжаем работать)  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyErrRezult = 1
        Me.Close()
    End Sub

    Private Sub ErrForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        RichTextBox1.Text = MyErrStr
    End Sub
End Class