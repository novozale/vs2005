Public Partial Class ItemsReturnQTY
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ƒобавление новой записи о кол-ве возвратов
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("SC01001").DefaultValue = InsertSC01001.Text
        SqlDataSource1.InsertParameters("QTY").DefaultValue = InsertQTY.Text
        SqlDataSource1.Insert()

        InsertSC01001.Text = ""
        InsertQTY.Text = ""
    End Sub
End Class