Public Partial Class RexelProductGroup
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой группы продуктов Rexel
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("RPGCode").DefaultValue = InsertRPGCode.Text
        SqlDataSource1.InsertParameters("RussianName").DefaultValue = InsertRussianName.Text
        SqlDataSource1.InsertParameters("EnglishName").DefaultValue = InsertEnglishName.Text
        SqlDataSource1.Insert()

        InsertRPGCode.Text = ""
        InsertRussianName.Text = ""
        InsertEnglishName.Text = ""
    End Sub
End Class