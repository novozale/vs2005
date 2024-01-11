Public Class LowMarginReason
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// Выход из формы с записью причины
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyMarginReason = TextBox1.Text
        Me.Close()
    End Sub
End Class
