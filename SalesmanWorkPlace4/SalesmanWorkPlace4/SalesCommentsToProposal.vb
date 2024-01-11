Public Class SalesCommentsToProposal

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с сохранением
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '----сохранение результатов
        If SaveRequest() = True Then
            Me.Close()
        End If
    End Sub

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных введенных в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Update tbl_SupplSearch_PropItems "
        MySQLStr = MySQLStr & "SET SalesmanComments = N'" & Trim(TextBox6.Text) & "' "
        MySQLStr = MySQLStr & "WHERE(ID = " & Label3.Text & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MySearchSupplier.DataGridView3.SelectedRows.Item(0).Cells("Comments").Value = Trim(TextBox6.Text)
        SaveRequest = True
    End Function

    Private Sub SalesCommentsToProposal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub SalesCommentsToProposal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка окна и данных в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Label3.Text = Declarations.MyItemPropID.ToString
        MySQLStr = "SELECT ID, ISNULL(SalesmanComments, '') AS SalesmanComments "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems "
        MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyItemPropID.ToString & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            TextBox6.Text = Declarations.MyRec.Fields("SalesmanComments").Value.ToString
        End If
        trycloseMyRec()
    End Sub
End Class