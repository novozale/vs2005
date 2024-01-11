Public Partial Class Cabel_MinSize
    Inherits System.Web.UI.Page


    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ƒобавление новой записи о кол-ве возвратов
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("ItemCode").DefaultValue = InsertItemCode.Text
        SqlDataSource1.InsertParameters("MinSize").DefaultValue = InsertMinSize.Text
        SqlDataSource1.Insert()

        InsertItemCode.Text = ""
        InsertMinSize.Text = ""
    End Sub
End Class