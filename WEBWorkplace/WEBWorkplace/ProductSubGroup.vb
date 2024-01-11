Public Class ProductSubGroup
    Public StartParam As String
    Public MySubGroupID As String

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
        Dim MyNewSubGroupCodeD As Double
        Dim MyNewSubGroupCode As String

        If CheckData() = True Then
            If StartParam = "Edit" Then '---обновление данных
                MySubGroupID = Trim(Declarations.MyProductGroupID) & Trim(Declarations.MyProductSubGroupID)
                MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
                MySQLStr = MySQLStr & "SET Name = N'" & Trim(TextBox3.Text) & "', "
                MySQLStr = MySQLStr & "Description = N'" & Trim(TextBox4.Text) & "', "
                MySQLStr = MySQLStr & "Rezerv1 = N'" & Trim(TextBox5.Text) & "', "
                MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, "
                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END "
                MySQLStr = MySQLStr & "WHERE (SubgroupID = N'" & MySubGroupID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else                        '---запись новых данных
                '---получение нового кода
                MySQLStr = "SELECT CONVERT(numeric, ISNULL(MAX(SubgroupCode), 0)) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
                MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Trim(Declarations.MyProductGroupID) & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    MyNewSubGroupCodeD = 0
                    trycloseMyRec()
                Else
                    MyNewSubGroupCodeD = Declarations.MyRec.Fields("CC").Value
                    trycloseMyRec()
                End If
                MyNewSubGroupCodeD = MyNewSubGroupCodeD + 1
                MyNewSubGroupCode = Microsoft.VisualBasic.Right("0000" & CStr(MyNewSubGroupCodeD), 4)
                Declarations.MyProductSubGroupID = MyNewSubGroupCode
                MySubGroupID = Trim(Declarations.MyProductGroupID) & MyNewSubGroupCode
                '---Запись нового значения
                Try
                    MySQLStr = "INSERT INTO tbl_WEB_ItemSubGroup "
                    MySQLStr = MySQLStr & "(SubgroupID, SubgroupCode, GroupCode, Name, Description, Rezerv1, RMStatus, WEBStatus) "
                    MySQLStr = MySQLStr & "VALUES (N'" & MySubGroupID & "', "
                    MySQLStr = MySQLStr & "N'" & MyNewSubGroupCode & "', "
                    MySQLStr = MySQLStr & "N'" & Trim(Declarations.MyProductGroupID) & "', "
                    MySQLStr = MySQLStr & "N'" & Trim(TextBox3.Text) & "', "
                    MySQLStr = MySQLStr & "N'" & Trim(TextBox4.Text) & "', "
                    MySQLStr = MySQLStr & "N'" & Trim(TextBox5.Text) & "', "
                    MySQLStr = MySQLStr & "1, "
                    MySQLStr = MySQLStr & "1) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex As Exception
                    MsgBox("Ошибка записи информации по новой группе  " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
            Me.Close()
        End If
    End Sub

    Private Sub ProductSubGroup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If StartParam = "Edit" Then
            MySQLStr = "SELECT Name, Description, Rezerv1 "
            MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemSubGroup WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Trim(Declarations.MyProductGroupID) & "') "
            MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & Trim(Declarations.MyProductSubGroupID) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Выделенная подгруппа товаров не найдена, возможно удален другим пользователем. закройте и откройте по новой вкладку подгрупп товаров.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Me.Close()
            Else
                TextBox1.Text = Trim(Declarations.MyProductGroupID)
                TextBox2.Text = Trim(Declarations.MyProductSubGroupID)
                TextBox3.Text = Declarations.MyRec.Fields("Name").Value
                TextBox4.Text = Declarations.MyRec.Fields("Description").Value
                TextBox5.Text = Declarations.MyRec.Fields("Rezerv1").Value
                trycloseMyRec()
            End If
        Else
            TextBox1.Text = Trim(Declarations.MyProductGroupID)
            TextBox2.Text = "N"
        End If
        TextBox1.Enabled = False
        TextBox2.Enabled = False
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""Название подгруппы продуктов"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            Exit Function
        End If

        CheckData = True
    End Function
End Class