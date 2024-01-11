Public Partial Class RelocationOrderR4
    Inherits System.Web.UI.Page

    Private Sub RelocationOrderR4_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение параметров страницы и их присвоение
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
        '// Подсветка строк
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
        '// Нажатие кнопки - Перенос заказа в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ChkAndTrsfToScala()
    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки - Перенос заказа в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ChkAndTrsfToScala()
    End Sub

    Protected Sub ChkAndTrsfToScala()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка правильности введения данных и перенос заказа в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            TransferToScala()
        End If

    End Sub

    Protected Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка правильности введения данных в форму
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLbl As Label                              'объект для получения свойств Label
        Dim MyTxt As TextBox                            'объект для получения свойств TextBox
        Dim Counter As Integer                          'счетчик
        Dim MyDCQTY As Double                           'доступное на DC количество
        Dim MyQTY As Double                             'заказанное количество
        Dim OrderQTY As Double                          'количество в заказе

        CheckData = True
        Label1.Text = ""
        OrderQTY = 0

        For Counter = 0 To GridView1.Rows.Count - 1
            '---свободно на DC
            MyLbl = GridView1.Rows(Counter).Cells(3).FindControl("FreeDC")
            If MyLbl.Text <> "" Then
                MyDCQTY = CDbl(MyLbl.Text)
            Else
                MyDCQTY = 0
            End If

            '---заказанное к перемещению количество
            MyTxt = GridView1.Rows(Counter).Cells(5).FindControl("QTY")
            MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
            If MyTxt.Text <> "" Then
                '---проверка - м.б. не число
                Try
                    MyQTY = CDbl(MyTxt.Text)
                Catch
                    Label1.Text = Label1.Text & "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число." & Chr(13) & Chr(10)
                    MyQTY = 0
                    CheckData = False
                End Try
                If InStr(MyTxt.Text, ",") > 0 Then
                    Label1.Text = Label1.Text & "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число." & Chr(13) & Chr(10)
                    MyQTY = 0
                    CheckData = False
                End If
            Else
                MyQTY = 0
            End If

            '---сверяем кол - во заказанного со свободным количеством на DC
            If MyQTY <> 0 And MyQTY > MyDCQTY Then
                Label1.Text = Label1.Text & "Код " & MyLbl.Text & "  Введенное количество больше, чем доступное на DC." & Chr(13) & Chr(10)
                MyQTY = MyDCQTY
                CheckData = False
            End If

            '---Проверяем - есть ли кабель, и если есть, выдаем сообщение
            If (Left(MyLbl.Text, 2) = "02" Or Left(MyLbl.Text, 2) = "03" Or Left(MyLbl.Text, 2) = "04" _
                Or Left(MyLbl.Text, 2) = "05" Or Left(MyLbl.Text, 2) = "06") And MyQTY <> 0 Then
                Label1.Text = Label1.Text & "Код " & MyLbl.Text & "  это кабель. Перемещается вручную." & "<br>"
                MyQTY = 0
                CheckData = False
            End If

            OrderQTY = OrderQTY + MyQTY
        Next

        If OrderQTY = 0 Then
            Label1.Text = "Общее количество запасов в закзае на перемещение равно 0. Такой заказ не будет сформирован." & vbCrLf
            CheckData = False
        End If

    End Function

    Protected Sub TransferToScala()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перенос данных в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim Conn As New OleDb.OleDbConnection("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")
        Dim MyLbl As Label                              'объект для получения свойств Label
        Dim MyTxt As TextBox                            'объект для получения свойств TextBox
        Dim Counter As Integer                          'счетчик
        Dim MyCode As String                            'код запаса
        Dim MyQTY As Double                             'заказанное количество
        Dim MyOrder As String                           'Номер заказа на перемещение в Scala

        '-------создание временной таблицы----------------------------------------------
        '---Удаление старой временной таблицы
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
        Dim objCmd As New OleDb.OleDbCommand(MySQLStr, Conn)
        Try
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            Label1.Text = "Ошибка N 1 процедуры переноса заказа в Scala. " & ex.Message
        End Try

        '---Создание новой временной таблицы
        MySQLStr = "CREATE TABLE #_MyOrder( "
        MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                 '--код товара в Scala
        MySQLStr = MySQLStr & "[QTY] decimal, "                             '--количество
        MySQLStr = MySQLStr & "[RestQTY] decimal  "                         '--Остаток - неперемещенное количество
        MySQLStr = MySQLStr & ") "
        Try
            objCmd.CommandText = MySQLStr
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            Label1.Text = "Ошибка N 2 процедуры переноса заказа в Scala. " & ex.Message
        End Try

        '-------Заполнение временной таблицы данными из формы----------------------------
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
                        Label1.Text = "Ошибка N 3 процедуры переноса заказа в Scala. " & ex.Message
                    End Try
                End If
            End If
        Next

        '--------------Запуск процедуры формирования заказа на перемещение---------------
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
        objCmd.Parameters("@MyOtherWHFlag").Value = 0 '--в заказ на перемещение запасы для заказов на продажу на других складах не включаем

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MyRelocOrderNum", OleDb.OleDbType.VarChar, 10))
        objCmd.Parameters("@MyRelocOrderNum").Direction = ParameterDirection.Output
        objCmd.Parameters("@MyRelocOrderNum").IsNullable = True

        Try
            objCmd.ExecuteNonQuery()
            MyOrder = objCmd.Parameters("@MyRelocOrderNum").Value
            Label6.Text = MyOrder
        Catch ex As Exception
            Label1.Text = "Ошибка N 4 процедуры переноса заказа в Scala. " & ex.Message
        End Try
        objCmd.Connection.Close()

        objCmd = Nothing
    End Sub
End Class