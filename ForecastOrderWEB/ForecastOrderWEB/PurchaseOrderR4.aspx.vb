Public Partial Class PurchaseOrder
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'ReCalcOrder()
    End Sub

    Private Sub PurchaseOrder_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение параметров страницы и их присвоение
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MySupCode As String
        Dim MyWarNo As String

        'If InStr(Request.ServerVariables("HTTP_REFERER"), "http://spbprd5/ReportServer") <> 1 And _
        'InStr(Request.ServerVariables("HTTP_REFERER"), "http://spbprd5/MD/PurchaseOrderR4.aspx") <> 1 Then
        'Response.Status = "301 Moved Permanently"
        'Response.AddHeader("Location", "http://spbprd5/reportServer")
        'End If

        'MySupCode = Request("MySupCode")
        'MyWarNo = Request("MyWarNo")
        MySupCode = "3432"
        MyWarNo = "01"

        Label3.Text = MySupCode
        Label5.Text = MyWarNo
    End Sub

    Private Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.DataItem("Price") = 0) Then
                e.Row.BackColor = Drawing.Color.LightPink
            Else
                If (e.Row.DataItem("RecQTY") <> 0) Then
                    e.Row.BackColor = Drawing.Color.LightGreen
                End If
            End If
        End If
    End Sub

    Protected Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и пересчет заказа
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ReCalcOrder()
        Dim MyLbl As Label                              'объект для получения свойств Label
        Dim MyTxt As TextBox                            'объект для получения свойств TextBox
        Dim Counter As Integer                          'счетчик
        Dim MyPrice As Double                           'цена запаса
        Dim MyQTY As Double                             'заказанное количество
        Dim OrderSum As Double                          'сумма заказа

        Label1.Text = ""
        OrderSum = 0
        For Counter = 0 To GridView1.Rows.Count - 1
            MyLbl = GridView1.Rows(Counter).Cells(3).FindControl("Price")
            If MyLbl.Text <> "" Then
                MyPrice = CDbl(MyLbl.Text)
            Else
                MyPrice = 0
            End If
            MyTxt = GridView1.Rows(Counter).Cells(7).FindControl("QTY")
            If MyTxt.Text <> "" Then
                MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
                '---проверка - м.б. не число
                Try
                    MyQTY = CDbl(MyTxt.Text)
                Catch
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyTxt.Text = ""
                    MyQTY = 0
                End Try
                If InStr(MyTxt.Text, ",") > 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyTxt.Text = ""
                    MyQTY = 0
                End If
                If MyPrice = 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Отсутствует закупочная цена и кол - во на закупку не равно 0."
                    MyTxt.Text = ""
                    MyQTY = 0
                End If
            Else
                MyQTY = 0
            End If
            OrderSum = OrderSum + MyQTY * MyPrice
        Next
        If GridView1.Rows.Count > 0 Then
            MyLbl = GridView1.Rows(0).Cells(4).FindControl("Curr")
            Label7.Text = OrderSum.ToString + " " + MyLbl.Text
        Else
            Label7.Text = OrderSum.ToString
        End If
    End Sub

    Protected Sub ReCalcOrder()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ф-ция Проверка и пересчет заказа
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLbl As Label                              'объект для получения свойств Label
        Dim MyTxt As TextBox                            'объект для получения свойств TextBox
        Dim Counter As Integer                          'счетчик
        Dim MyPrice As Double                           'цена запаса
        Dim MyQTY As Double                             'заказанное количество
        Dim OrderSum As Double                          'сумма заказа

        Label1.Text = ""
        OrderSum = 0
        For Counter = 0 To GridView1.Rows.Count - 1
            MyLbl = GridView1.Rows(Counter).Cells(3).FindControl("Price")
            If MyLbl.Text <> "" Then
                MyPrice = CDbl(MyLbl.Text)
            Else
                MyPrice = 0
            End If
            MyTxt = GridView1.Rows(Counter).Cells(7).FindControl("QTY")
            If MyTxt.Text <> "" Then
                MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
                '---проверка - м.б. не число
                Try
                    MyQTY = CDbl(MyTxt.Text)
                Catch
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyTxt.Text = ""
                    MyQTY = 0
                End Try
                If InStr(MyTxt.Text, ",") > 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyTxt.Text = ""
                    MyQTY = 0
                End If
                If MyPrice = 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Отсутствует закупочная цена и кол - во на закупку не равно 0."
                    MyTxt.Text = ""
                    MyQTY = 0
                End If
            Else
                MyQTY = 0
            End If
            OrderSum = OrderSum + MyQTY * MyPrice
        Next
        If GridView1.Rows.Count > 0 Then
            MyLbl = GridView1.Rows(0).Cells(4).FindControl("Curr")
            Label7.Text = OrderSum.ToString + " " + MyLbl.Text
        Else
            Label7.Text = OrderSum.ToString
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки - Перенос заказа в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////

        ChkAndTrsfToScala()
    End Sub

    Protected Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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

        If CheckData = True Then
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
        Dim MyPrice As Double                           'цена запаса
        Dim MyQTY As Double                             'заказанное количество
        Dim OrderSum As Double                          'сумма заказа

        CheckData = True
        Label1.Text = ""
        OrderSum = 0
        For Counter = 0 To GridView1.Rows.Count - 1
            MyLbl = GridView1.Rows(Counter).Cells(3).FindControl("Price")
            If MyLbl.Text <> "" Then
                MyPrice = CDbl(MyLbl.Text)
            Else
                MyPrice = 0
            End If
            MyTxt = GridView1.Rows(Counter).Cells(7).FindControl("QTY")
            If MyTxt.Text <> "" Then
                MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
                '---проверка - м.б. не число
                Try
                    MyQTY = CDbl(MyTxt.Text)
                Catch
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyTxt.Text = ""
                    MyQTY = 0
                    CheckData = False
                End Try
                If InStr(MyTxt.Text, ",") > 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Введено неверное количество. Должно быть число."
                    MyQTY = 0
                    CheckData = False
                End If
                If MyPrice = 0 Then
                    Label1.Text = "Код " & MyLbl.Text & "  Отсутствует закупочная цена и кол - во на закупку не равно 0."
                    MyTxt.Text = ""
                    MyQTY = 0
                    CheckData = False
                End If
            Else
                MyQTY = 0
            End If
            OrderSum = OrderSum + MyQTY * MyPrice
        Next
        If OrderSum = 0 Then
            CheckData = False
        End If
        If GridView1.Rows.Count > 0 Then
            MyLbl = GridView1.Rows(0).Cells(4).FindControl("Curr")
            Label7.Text = OrderSum.ToString + " " + MyLbl.Text
        Else
            Label7.Text = OrderSum.ToString
        End If
    End Function

    Protected Sub TransferToScala()
        '/////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перенос данных в Scala
        '//
        '/////////////////////////////////////////////////////////////////////////////////////////
        Dim strSQL As String
        Dim Conn As New OleDb.OleDbConnection("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")
        'Dim Conn As New OleDb.OleDbConnection("Provider=SQLOLEDB.1;Server=spbdvl2;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")
        Dim MyLbl As Label                              'объект для получения свойств Label
        Dim MyTxt As TextBox                            'объект для получения свойств TextBox
        Dim Counter As Integer                          'счетчик
        Dim MyCode As String                            'цена запаса
        Dim MyQTY As Double                             'заказанное количество
        Dim MyOrder As String                           'Номер заказа на закупку в Scala

        '-------Создание заголовка заказа-----------------------------------------------
        strSQL = "spp_ForecastOrderR4_PurchaseOrder_CreateHeader"
        Dim objCmd As New OleDb.OleDbCommand(strSQL, Conn)
        objCmd.CommandTimeout = 600
        objCmd.CommandType = CommandType.StoredProcedure

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MySupCode", OleDb.OleDbType.VarChar, 50))
        objCmd.Parameters("@MySupCode").Direction = ParameterDirection.Input
        objCmd.Parameters("@MySupCode").Value = Label3.Text

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MyWarNo", OleDb.OleDbType.VarChar, 6))
        objCmd.Parameters("@MyWarNo").Direction = ParameterDirection.Input
        objCmd.Parameters("@MyWarNo").Value = Label5.Text

        objCmd.Parameters.Add(New OleDb.OleDbParameter("@MyOrderNumRet", OleDb.OleDbType.VarChar, 10))
        objCmd.Parameters("@MyOrderNumRet").Direction = ParameterDirection.Output
        objCmd.Parameters("@MyOrderNumRet").IsNullable = True

        Try
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
            MyOrder = objCmd.Parameters("@MyOrderNumRet").Value
            Label6.Text = MyOrder
        Catch ex As Exception
            Label1.Text = "Ошибка запуска процедуры переноса заказа в Scala. " & ex.Message
        End Try
        objCmd.Connection.Close()
        objCmd = Nothing

        '-------Создание строк заказа---------------------------------------------------
        For Counter = 0 To GridView1.Rows.Count - 1
            MyLbl = GridView1.Rows(Counter).Cells(0).FindControl("Code")
            If MyLbl.Text <> "" Then
                MyCode = MyLbl.Text
                MyTxt = GridView1.Rows(Counter).Cells(7).FindControl("QTY")
                If Trim(MyTxt.Text) <> "" And CDbl(MyTxt.Text) <> 0 Then
                    MyQTY = CDbl(MyTxt.Text)

                    strSQL = "spp_ForecastOrderR4_PurchaseOrder_CreateRow"
                    Dim objCmd1 As New OleDb.OleDbCommand(strSQL, Conn)
                    objCmd1.CommandTimeout = 600
                    objCmd1.CommandType = CommandType.StoredProcedure

                    objCmd1.Parameters.Add(New OleDb.OleDbParameter("@MyOrderNum", OleDb.OleDbType.VarChar, 10))
                    objCmd1.Parameters("@MyOrderNum").Direction = ParameterDirection.Input
                    objCmd1.Parameters("@MyOrderNum").Value = Label6.Text

                    objCmd1.Parameters.Add(New OleDb.OleDbParameter("@MyItemCode", OleDb.OleDbType.VarChar, 35))
                    objCmd1.Parameters("@MyItemCode").Direction = ParameterDirection.Input
                    objCmd1.Parameters("@MyItemCode").Value = MyCode

                    objCmd1.Parameters.Add(New OleDb.OleDbParameter("@MyQTY", OleDb.OleDbType.Double))
                    objCmd1.Parameters("@MyQTY").Direction = ParameterDirection.Input
                    objCmd1.Parameters("@MyQTY").Value = MyQTY

                    Try
                        objCmd1.Connection.Open()
                        objCmd1.ExecuteNonQuery()
                    Catch ex As Exception
                        Label1.Text = "Ошибка запуска процедуры переноса заказа в Scala. " & ex.Message
                    End Try
                    objCmd1.Connection.Close()
                    objCmd1 = Nothing

                End If
            End If
        Next


    End Sub

    Private Sub SqlDataSource1_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles SqlDataSource1.Selecting
        e.Command.CommandTimeout = 600
    End Sub
End Class