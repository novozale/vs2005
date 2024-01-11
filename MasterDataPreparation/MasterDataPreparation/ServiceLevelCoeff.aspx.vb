Public Partial Class ServiceLevelCoeff
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обработка нажатия кнопки
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        Dim strSQL As String = "spp_ForecastOrderR4_Main_DC"
        Dim strSQL1 As String = "spp_ForecastOrderR4_Main_RWH"
        ExecuteStatement(strSQL, strSQL1)
    End Sub

    Function ExecuteStatement(ByVal strSQL, ByVal strSQL1)
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запуск хранимой процедуры для пересчета ROP, МЖЗ, ABC, XYZ
        '//
        '////////////////////////////////////////////////////////////////////////////////////
        Dim Conn As New OleDb.OleDbConnection("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")

        Dim objCmd As New OleDb.OleDbCommand(strSQL, Conn)
        objCmd.CommandTimeout = 2000
        objCmd.CommandType = CommandType.StoredProcedure

        Dim objCmd1 As New OleDb.OleDbCommand(strSQL1, Conn)
        objCmd1.CommandTimeout = 2000
        objCmd1.CommandType = CommandType.StoredProcedure

        Try
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
            objCmd1.Connection.Open()
            objCmd1.ExecuteNonQuery()
            Label1.Text = "ROP и МЖЗ пересчитаны."
        Catch ex As Exception
            Label1.Text = "Ошибка запуска процедуры пересчета."
        End Try

        objCmd.Connection.Close()
        objCmd1.Connection.Close()
    End Function
End Class