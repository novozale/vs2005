Public Partial Class MarginCoeff
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////////

        Dim strSQL As String = "spp_PrepareCommonPriceList_PriCost"
        ExecuteStatement(strSQL)
    End Sub

    Function ExecuteStatement(ByVal strSQL)
        '////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ��������� ��� ��������� ROP, ���, ABC, XYZ
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
            Label1.Text = "����� - ���� �������� � ��������� ��������������"
        Catch ex As Exception
            Label1.Text = "������ ������� ��������� ��������� ����� - �����."
        End Try

        objCmd.Connection.Close()
    End Function
End Class