Public Partial Class Correctures
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой записи об исключениях / корректировке
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("OperTip").DefaultValue = InsertOperTip.Text
        SqlDataSource1.InsertParameters("SupCode").DefaultValue = InsertSupCode.Text
        SqlDataSource1.InsertParameters("Supplier").DefaultValue = InsertSupplier.Text
        SqlDataSource1.InsertParameters("CustCode").DefaultValue = InsertCustCode.Text
        SqlDataSource1.InsertParameters("Customer").DefaultValue = InsertCustomer.Text
        SqlDataSource1.InsertParameters("Sales").DefaultValue = InsertSales.Text
        SqlDataSource1.InsertParameters("Cost").DefaultValue = InsertCost.Text
        SqlDataSource1.InsertParameters("InsDate").DefaultValue = InsertInsDate.Text
        SqlDataSource1.Insert()

        InsertOperTip.Text = "0"
        InsertSupCode.Text = ""
        InsertSupplier.Text = ""
        InsertCustCode.Text = ""
        InsertCustomer.Text = ""
        InsertSales.Text = ""
        InsertCost.Text = ""
        InsertInsDate.Text = CStr(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))
    End Sub
End Class