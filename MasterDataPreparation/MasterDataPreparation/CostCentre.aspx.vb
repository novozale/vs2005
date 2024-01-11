Public Partial Class CostCentre
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление нового кост центра
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("CCNo").DefaultValue = InsertCCNo.Text
        SqlDataSource1.InsertParameters("City").DefaultValue = InsertCity.Text
        SqlDataSource1.InsertParameters("CCName").DefaultValue = InsertCCName.Text
        SqlDataSource1.InsertParameters("B2B").DefaultValue = InsertB2B.Text
        SqlDataSource1.InsertParameters("Retail").DefaultValue = InsertRetail.Text
        SqlDataSource1.Insert()

        InsertCCNo.Text = ""
        InsertCity.Text = ""
        InsertCCName.Text = ""
        InsertB2B.Text = "Да"
        InsertRetail.Text = "Да"
    End Sub
End Class