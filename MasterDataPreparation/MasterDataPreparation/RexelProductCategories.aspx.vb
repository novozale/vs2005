Public Partial Class RexelProductCategories
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ƒобавление новой категории продуктов Rexel
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("CategoryNum").DefaultValue = InsertCategoryNum.Text
        SqlDataSource1.InsertParameters("CategoryName").DefaultValue = InsertCategoryName.Text
        SqlDataSource1.InsertParameters("RPGCode").DefaultValue = InsertRPGCode.SelectedValue
        SqlDataSource1.InsertParameters("SRPGCode").DefaultValue = InsertSRPGCode.SelectedValue
        SqlDataSource1.Insert()

        InsertCategoryNum.Text = ""
        InsertCategoryName.Text = ""
    End Sub
End Class