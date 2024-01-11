Public Class Product
    Public StartParam As String
    Public MyGroup As String

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

        MySQLStr = "UPDATE tbl_WEB_Items "
        MySQLStr = MySQLStr & "SET WEBName = N'" & Trim(TextBox3.Text) & "', "
        MySQLStr = MySQLStr & "SubGroupCode = N'" & ComboBox1.SelectedValue & "', "
        MySQLStr = MySQLStr & "Description = N'" & Trim(TextBox11.Text) & "', "
        MySQLStr = MySQLStr & "Rezerv = N'" & Trim(TextBox13.Text) & "', "
        MySQLStr = MySQLStr & "RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE CASE WHEN RMStatus = 2 THEN 2 ELSE 3 END END, "
        MySQLStr = MySQLStr & "WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE CASE WHEN WEBStatus = 2 THEN 2 ELSE 3 END END "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(Declarations.MyProductID) & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        Me.Close()
    End Sub

    Private Sub Product_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка подгрупп
        Dim MyDs As New DataSet

        '---------------Список подгрупп
        MySQLStr = "SELECT SubgroupCode, LTRIM(RTRIM(SubgroupCode)) + ' ' + LTRIM(RTRIM(Name)) AS Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup "
        MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Trim(MyGroup) & "') "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT '' AS SubgroupCode, '' AS Name "
        MySQLStr = MySQLStr & "ORDER BY SubgroupCode "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "SubgroupCode"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '-------------Загрузка значений
        MySQLStr = "SELECT tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Items.WEBName, tbl_WEB_Items.ManufacturerCode, ISNULL(tbl_WEB_Manufacturers.Name, "
        MySQLStr = MySQLStr & "'') AS ManufacturerName, tbl_WEB_Items.CountryCode, tbl_WEB_Items.Country, tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.GroupCode, "
        MySQLStr = MySQLStr & "ISNULL(tbl_WEB_ItemGroup.Name, '') AS GroupName, tbl_WEB_Items.SubGroupCode, tbl_WEB_Items.Description, tbl_WEB_Items.WHAssortiment, "
        MySQLStr = MySQLStr & "tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, tbl_WEB_Pictures.Picture "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemGroup.Code LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.Code = N'" & Trim(Declarations.MyProductID) & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Выделенная подгруппа товаров не найдена, возможно удален другим пользователем. закройте и откройте по новой вкладку подгрупп товаров.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Me.Close()
        Else
            TextBox1.Text = Declarations.MyRec.Fields("Code").Value
            TextBox2.Text = Declarations.MyRec.Fields("Name").Value
            TextBox3.Text = Declarations.MyRec.Fields("WEBName").Value
            TextBox4.Text = Declarations.MyRec.Fields("ManufacturerCode").Value
            TextBox5.Text = Declarations.MyRec.Fields("ManufacturerName").Value
            TextBox6.Text = Declarations.MyRec.Fields("CountryCode").Value
            TextBox7.Text = Declarations.MyRec.Fields("Country").Value
            TextBox8.Text = Declarations.MyRec.Fields("ManufacturerItemCode").Value
            TextBox9.Text = Declarations.MyRec.Fields("GroupCode").Value
            TextBox10.Text = Declarations.MyRec.Fields("GroupName").Value
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("SubGroupCode").Value
            TextBox11.Text = Declarations.MyRec.Fields("Description").Value
            If Declarations.MyRec.Fields("WHAssortiment").Value = 0 Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
            TextBox12.Text = Declarations.MyRec.Fields("UOM").Value
            TextBox13.Text = Declarations.MyRec.Fields("Rezerv").Value
            If Declarations.MyRec.Fields("Picture").Value.Equals(System.DBNull.Value) Then
            Else
                Dim ms As New IO.MemoryStream(CType(Declarations.MyRec.Fields("Picture").Value, Byte()))
                Dim picture As Image

                Try
                    picture = Image.FromStream(ms)
                    PictureBox1.Image = picture
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
            trycloseMyRec()
        End If
        CheckButtons()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление картинки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyAnswer As Object
        Dim MySQLStr As String

        MyAnswer = MsgBox("Вы уверены, что хотите удалить картинку?", MsgBoxStyle.YesNo, "Внимание!")
        If MyAnswer = DialogResult.Yes Then
            MySQLStr = "DELETE FROM tbl_WEB_Pictures "
            MySQLStr = MySQLStr & "WHERE (ScalaItemCode = N'" & Trim(TextBox1.Text) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            PictureBox1.Image = Nothing
            CheckButtons()
        End If
    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If IsNothing(PictureBox1.Image) = True Then
            Button3.Enabled = False
        Else
            Button3.Enabled = True
        End If
    End Sub
End Class