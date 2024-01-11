Public Partial Class CountryNameAndID
    Inherits System.Web.UI.Page

   

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ƒобавление новой записи о стране - производителе
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("Name").DefaultValue = InsertName.Text
        SqlDataSource1.InsertParameters("Code").DefaultValue = InsertCode.Text
        SqlDataSource1.Insert()

        InsertName.Text = ""
        InsertCode.Text = ""
    End Sub
End Class