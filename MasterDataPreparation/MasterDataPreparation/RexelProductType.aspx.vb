Public Partial Class RexelProductType
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление нового типа продуктов Rexel
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("SRPGCode").DefaultValue = InsertSRPGCode.Text
        SqlDataSource1.InsertParameters("RussianName").DefaultValue = InsertRussianName.Text
        SqlDataSource1.InsertParameters("EnglishName").DefaultValue = InsertEnglishName.Text
        SqlDataSource1.Insert()

        InsertSRPGCode.Text = ""
        InsertRussianName.Text = ""
        InsertEnglishName.Text = ""
    End Sub
End Class