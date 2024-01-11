Public Partial Class SalesInvoicesClaims
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой записи о претензии
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("InsertSalesInvoiceNum").DefaultValue = InsertSalesInvoiceNum.Text
        SqlDataSource1.Insert()

        InsertSalesInvoiceNum.Text = ""
    End Sub
End Class