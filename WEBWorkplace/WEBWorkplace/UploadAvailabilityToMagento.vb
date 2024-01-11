Public Class UploadAvailabilityToMagento

    Private Sub UploadAvailabilityToMagento_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
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
        UploadAvailability_ToMagento()
        Me.Cursor = Cursors.Default
        Button2.Enabled = True
        Button3.Enabled = True
        MyErrWindow = New ErrWindow
        MyErrWindow.ShowDialog()
    End Sub
End Class