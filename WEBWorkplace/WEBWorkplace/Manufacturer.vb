Public Class Manufacturer

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(TextBox3.Text) & "', "
        MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(TextBox4.Text) & "', "
        MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, "
        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END "
        MySQLStr = MySQLStr & "WHERE (ID = " & CStr(Declarations.MyManufacturerID) & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        Me.Close()
    End Sub

    Private Sub Manufacturer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "SELECT ID, Name, WEBName, Rezerv1 "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Manufacturers WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = " & UCase(Trim(Declarations.MyManufacturerID)) & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Выделенный производитель не найден, возможно удален другим пользователем. закройте и откройте по новой вкладку производителей.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Me.Close()
        Else
            TextBox1.Text = Declarations.MyRec.Fields("ID").Value.ToString
            TextBox2.Text = Declarations.MyRec.Fields("Name").Value
            TextBox3.Text = Declarations.MyRec.Fields("WEBName").Value.ToString
            TextBox4.Text = Declarations.MyRec.Fields("Rezerv1").Value
            trycloseMyRec()
        End If
        TextBox1.Enabled = False
        TextBox2.Enabled = False
    End Sub
End Class