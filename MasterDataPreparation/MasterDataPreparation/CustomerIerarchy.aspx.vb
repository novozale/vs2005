Public Partial Class CustomerIerarchy
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой записи о группировке клиентов
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("CustomerCode").DefaultValue = InsertCustomerCode.Text
        SqlDataSource1.InsertParameters("CustomerName").DefaultValue = InsertCustomerName.Text
        SqlDataSource1.InsertParameters("JoinCode").DefaultValue = InsertJoinCode.Text
        SqlDataSource1.InsertParameters("Flag").DefaultValue = InsertFlag.Text
        SqlDataSource1.InsertParameters("Vid").DefaultValue = InsertVid.Text
        SqlDataSource1.Insert()

        InsertCustomerCode.Text = ""
        InsertCustomerName.Text = ""
        InsertJoinCode.Text = ""
        InsertFlag.Text = "1"
        InsertVid.Text = "Customer"
    End Sub

    Private Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка срок с информацией по клиентам 2 уровня (обобщенным)
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.DataItem("Flag") = 2) Then
                e.Row.BackColor = Drawing.Color.LightGray
            End If
        End If
    End Sub
End Class