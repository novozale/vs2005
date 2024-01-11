Public Partial Class MarginCoeffByItem
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление новой записи по запасу с указанием коэфф маржи
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        SqlDataSource1.InsertParameters("SC01001").DefaultValue = InsertSC01001.Text
        SqlDataSource1.InsertParameters("MarginCoeff").DefaultValue = InsertMarginCoeff.Text
        SqlDataSource1.Insert()

        InsertSC01001.Text = ""
        InsertMarginCoeff.Text = ""
    End Sub

    Protected Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обработка нажатия кнопки
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        Dim strSQL As String = "spp_PrepareCommonPriceList_PriCost"
        ExecuteStatement(strSQL)
    End Sub

    Function ExecuteStatement(ByVal strSQL)
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запуск хранимой процедуры для пересчета ROP, МЖЗ, ABC, XYZ
        '//
        '////////////////////////////////////////////////////////////////////////////////////
        Dim Conn As New OleDb.OleDbConnection _
                 ("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")

        Dim objCmd As New OleDb.OleDbCommand(strSQL, Conn)
        objCmd.CommandTimeout = 600
        objCmd.CommandType = CommandType.StoredProcedure
        Try
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
            Label1.Text = "Прайс - лист расчитан с заданными коэффициентами"
        Catch ex As Exception
            Label1.Text = "Ошибка запуска процедуры пересчета прайс - листа."
        End Try

        objCmd.Connection.Close()
    End Function
End Class