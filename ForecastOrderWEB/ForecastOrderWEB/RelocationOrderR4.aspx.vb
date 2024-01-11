Public Partial Class RelocationOrderR4
    Inherits System.Web.UI.Page

    Private Sub RelocationOrderR4_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �������� � �� ����������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MySrcWH As String
        Dim MyWarNo As String

        If InStr(Request.ServerVariables("HTTP_REFERER"), "http://spbprd5/ReportServer") <> 1 And _
            InStr(Request.ServerVariables("HTTP_REFERER"), "http://spbprd5/MD/RelocationOrderR4.aspx") <> 1 Then
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "http://spbprd5/reportServer")
        End If


        MyWarNo = Request("MyWarNo")
        MySrcWH = Request("MySrcWH")


        Label3.Text = MyWarNo
        Label5.Text = MySrcWH
    End Sub

    Private Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �����
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.DataItem("RecQTY") > e.Row.DataItem("FreeDC")) Then
                e.Row.BackColor = Drawing.Color.LightPink
            Else
                If (e.Row.DataItem("RecQTY") <> 0) Then
                    e.Row.BackColor = Drawing.Color.LightGreen
                End If
            End If
        End If
    End Sub

    Protected Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ - ������� ������ � Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ChkAndTrsfToScala()
    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ - ������� ������ � Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ChkAndTrsfToScala()
    End Sub

    Protected Sub ChkAndTrsfToScala()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ �������� ������ � ������� ������ � Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            TransferToScala()
        End If

    End Sub

    Protected Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ �������� ������ � �����
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLbl As Label                              '������ ��� ��������� ������� Label
        Dim MyTxt As TextBox                            '������ ��� ��������� ������� TextBox
        Dim Counter As Integer                          '�������
        Dim MyDCQTY As Double                           '��������� �� DC ����������
        Dim MyQTY As Double                             '���������� ����������
        Dim OrderQTY As Double                          '���������� � ������

        CheckData = True
        Label1.Text = ""
        OrderQTY = 0

        For Counter = 0 To GridView1.Rows.Count - 1
            '---�������� �� DC
            MyLbl = GridView1.Rows(Counter).Cells(3).FindControl("FreeDC")
            If MyLbl.Text <> "" Then
                MyDCQTY = CDbl(MyLbl.Text)
            Else
                MyDCQTY = 0
            End If

            '---���������� � ����������� ����������
            MyTxt = GridView1.Rows(Counter).Cells(5).FindControl("QTY")
            MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
            If MyTxt.Text <> "" Then
                '---�������� - �.�. �� �����
                Try
                    MyQTY = CDbl(MyTxt.Text)
                Catch
                    Label1.Text = Label1.Text & "��� " & MyLbl.Text & "  ������� �������� ����������. ������ ���� �����." & Chr(13) & Chr(10)
                    MyQTY = 0
                    CheckData = False
                End Try
                If InStr(MyTxt.Text, ",") > 0 Then
                    Label1.Text = Label1.Text & "��� " & MyLbl.Text & "  ������� �������� ����������. ������ ���� �����." & Chr(13) & Chr(10)
                    MyQTY = 0
                    CheckData = False
                End If
            Else
                MyQTY = 0
            End If

            '---������� ��� - �� ����������� �� ��������� ����������� �� DC
            If MyQTY <> 0 And MyQTY > MyDCQTY Then
                Label1.Text = Label1.Text & "��� " & MyLbl.Text & "  ��������� ���������� ������, ��� ��������� �� DC." & Chr(13) & Chr(10)
                MyQTY = MyDCQTY
                CheckData = False
            End If

            '---��������� - ���� �� ������, � ���� ����, ������ ���������
            If (Left(MyLbl.Text, 2) = "02" Or Left(MyLbl.Text, 2) = "03" Or Left(MyLbl.Text, 2) = "04" _
                Or Left(MyLbl.Text, 2) = "05" Or Left(MyLbl.Text, 2) = "06") And MyQTY <> 0 Then
                Label1.Text = Label1.Text & "��� " & MyLbl.Text & "  ��� ������. ������������ �������." & "<br>"
                MyQTY = 0
                CheckData = False
            End If

            OrderQTY = OrderQTY + MyQTY
        Next

        If OrderQTY = 0 Then
            Label1.Text = "����� ���������� ������� � ������ �� ����������� ����� 0. ����� ����� �� ����� �����������." & vbCrLf
            CheckData = False
        End If

    End Function

    Protected Sub TransferToScala()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ � Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim Conn As New OleDb.OleDbConnection("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")
        Dim MyLbl As Label                              '������ ��� ��������� ������� Label
        Dim MyTxt As TextBox                            '������ ��� ��������� ������� TextBox
        Dim Counter As Integer                          '�������
        Dim MyCode As String                            '��� ������
        Dim MyQTY As Double                             '���������� ����������
        Dim MyOrder As String                           '����� ������ �� ����������� � Scala

        '-------�������� ��������� �������----------------------------------------------
        '---�������� ������ ��������� �������
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
        Dim objCmd As New OleDb.OleDbCommand(MySQLStr, Conn)
        Try
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            Label1.Text = "������ N 1 ��������� �������� ������ � Scala. " & ex.Message
        End Try

        '---�������� ����� ��������� �������
        MySQLStr = "CREATE TABLE #_MyOrder( "
        MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                 '--��� ������ � Scala
        MySQLStr = MySQLStr & "[QTY] decimal, "                             '--����������
        MySQLStr = MySQLStr & "[RestQTY] decimal  "                         '--������� - �������������� ����������
        MySQLStr = MySQLStr & ") "
        Try
            objCmd.CommandText = MySQLStr
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            Label1.Text = "������ N 2 ��������� �������� ������ � Scala. " & ex.Message
        End Try

        '-------���������� ��������� ������� ������� �� �����----------------------------
        For Counter = 0 To GridView1.Rows.Count - 1
            MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
            If MyLbl.Text <> "" Then
                MyCode = MyLbl.Text
                MyTxt = GridView1.Rows(Counter).Cells(5).FindControl("QTY")
                If MyTxt.Text <> "" Then
                    MyQTY = CDbl(MyTxt.Text)
                    MySQLStr = "INSERT INTO #_MyOrder "
                    MySQLStr = MySQLStr & "(ItemCode, QTY, RestQTY) "
                    MySQLStr = MySQLStr & "VALUES (N'" & MyCode & "', "
                    MySQLStr = MySQLStr & CStr(MyQTY) & ", "
                    MySQLStr = MySQLStr & CStr(MyQTY) & ") "
                    objCmd.CommandText = MySQLStr
                    Try
                        objCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        Label1.Text = "������ N 3 ��������� �������� ������ � Scala. " & ex.Message
                    End Try
                End If
            End If
        Next

        '--------------������ ��������� ������������ ������ �� �����������---------------
        MySQLStr = "spp_ForecastOrderR4_RelocationOrder_Create"
        objCmd.CommandText = MySQLStr
        objCmd.CommandTimeout = 600
        objCmd.CommandType = CommandType.StoredProcedure

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@SrcWarNo", OleDb.OleDbType.VarChar, 6))
        objCmd.Parameters("@SrcWarNo").Direction = ParameterDirection.Input
        objCmd.Parameters("@SrcWarNo").Value = Label5.Text

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@DestWarNo", OleDb.OleDbType.VarChar, 6))
        objCmd.Parameters("@DestWarNo").Direction = ParameterDirection.Input
        objCmd.Parameters("@DestWarNo").Value = Label3.Text

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MyOtherWHFlag", OleDb.OleDbType.Integer))
        objCmd.Parameters("@MyOtherWHFlag").Direction = ParameterDirection.Input
        objCmd.Parameters("@MyOtherWHFlag").Value = 0 '--� ����� �� ����������� ������ ��� ������� �� ������� �� ������ ������� �� ��������

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MyRelocOrderNum", OleDb.OleDbType.VarChar, 10))
        objCmd.Parameters("@MyRelocOrderNum").Direction = ParameterDirection.Output
        objCmd.Parameters("@MyRelocOrderNum").IsNullable = True

        Try
            objCmd.ExecuteNonQuery()
            MyOrder = objCmd.Parameters("@MyRelocOrderNum").Value
            Label6.Text = MyOrder
        Catch ex As Exception
            Label1.Text = "������ N 4 ��������� �������� ������ � Scala. " & ex.Message
        End Try
        objCmd.Connection.Close()

        objCmd = Nothing
    End Sub
End Class