Public Partial Class RexelEndMarket
    Inherits System.Web.UI.Page
    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление нового типа рынка Rexel
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("EMCode").DefaultValue = InsertEMCode.Text
        SqlDataSource1.InsertParameters("RussianName").DefaultValue = InsertRussianName.Text
        SqlDataSource1.InsertParameters("EnglishName").DefaultValue = InsertEnglishName.Text
        SqlDataSource1.Insert()

        InsertEMCode.Text = ""
        InsertRussianName.Text = ""
        InsertEnglishName.Text = ""
    End Sub
End Class