
Public Class MainForm
    Public MyBS As New BindingSource()

    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отрисовка подписей к Tab горизонтально
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim g As Graphics = e.Graphics
        Dim _TextBrush As Brush
        Dim CustColor As Color
        Dim br As Brush

        '  Get the item from the collection. 
        Dim _TabPage As TabPage = TabControl1.TabPages(e.Index)

        '  Get the real bounds for the tab rectangle. 
        Dim _TabBounds As Rectangle = TabControl1.GetTabRect(e.Index)

        If (e.State = DrawItemState.Selected) Then
            '  Draw a different background color, and don't paint a focus rectangle.
            _TextBrush = New SolidBrush(Color.Black)
            CustColor = Color.FromKnownColor(KnownColor.ButtonFace) ' .FromArgb(100, 100, 100)
            br = New SolidBrush(CustColor)
            g.FillRectangle(br, e.Bounds)
        Else
            _TextBrush = New System.Drawing.SolidBrush(Color.Black)
            'e.DrawBackground()
            CustColor = Color.FromArgb(170, 170, 170) 'FromKnownColor(KnownColor.ButtonHighlight)  
            br = New SolidBrush(CustColor)
            g.FillRectangle(br, e.Bounds)
        End If

        '  Use our own font. 
        Dim _TabFont As New Font("Microsoft Sans Serife", 8.0, FontStyle.Bold, GraphicsUnit.Point)

        '  Draw string. Center the text. 
        Dim _StringFlags As New StringFormat()
        _StringFlags.Alignment = StringAlignment.Center
        _StringFlags.LineAlignment = StringAlignment.Center
        g.DrawString(_TabPage.Text, _TabFont, _TextBrush, _TabBounds, New StringFormat(_StringFlags))

    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим данные
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор закладки
        '// для выбранной закладки выводим данные
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        Select Case sender.selectedtab.text
            Case "Города"
                LoadCities()
                CheckCitiesButtons()
            Case "Производители"
                LoadManufacturers()
                CheckManufacturersButtons()
            Case "Продавцы"
                LoadSalesmans()
                CheckSalesmansButtons()
            Case "Группы товара"
                LoadProductGroup()
                CheckProductGroupButtons()
            Case "Подгруппы товара"
                LoadProductSubgroup()
                CheckProductSubGroupButtons()
            Case "Товары"
                LoadProducts()
                CheckProductButtons()
            Case "Клиенты"
                LoadCustomers()
                CheckCustomerButtons()
            Case "Скидки и соглас. ассортимент"
                LoadDiscountsHeader()
                'Case "Операции"
            Case Else
        End Select
    End Sub

    Private Sub LoadCities()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка городов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Cities "
        MySQLStr = MySQLStr & "ORDER BY ID "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Код города"
        DataGridView1.Columns(0).Width = 200
        DataGridView1.Columns(1).HeaderText = "Название города"
        DataGridView1.Columns(1).Width = 500

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadManufacturers()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка производителей
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ID, Name, WEBName, Rezerv1, RMStatus, WEBStatus "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "ORDER BY ID "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView2.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView2.Columns(0).HeaderText = "Код производителя"
        DataGridView2.Columns(0).Width = 100
        DataGridView2.Columns(1).HeaderText = "Название производителя"
        DataGridView2.Columns(1).Width = 290
        DataGridView2.Columns(2).HeaderText = "Корректное название производителя"
        DataGridView2.Columns(2).Width = 290
        DataGridView2.Columns(3).HeaderText = "Резервное поле"
        DataGridView2.Columns(3).Width = 250
        DataGridView2.Columns(4).HeaderText = "Статус Scala"
        DataGridView2.Columns(4).Width = 50
        DataGridView2.Columns(5).HeaderText = "Статус WEB"
        DataGridView2.Columns(5).Width = 50

        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadSalesmans()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка продавцов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_WEB_Salesmans.Code, tbl_WEB_Salesmans.Name, tbl_WEB_Salesmans.Email, ISNULL(tbl_WEB_Cities.Name, '') AS City, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_WEB_Salesmans.OfficeLeader = 0 THEN '' ELSE 'да' END AS OfficeLeader, CASE WHEN tbl_WEB_Salesmans.OnDuty = 0 THEN '' ELSE 'да' END AS OnDuty, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_WEB_Salesmans.IsActive = 0 THEN '' ELSE 'активный' END AS IsActive, tbl_WEB_Salesmans.Rezerv1, "
        MySQLStr = MySQLStr & "tbl_WEB_Salesmans.Rezerv2, tbl_WEB_Salesmans.RMStatus, tbl_WEB_Salesmans.WEBStatus, tbl_WEB_Salesmans.ScalaStatus "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Salesmans LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Cities ON tbl_WEB_Salesmans.City = tbl_WEB_Cities.ID "
        MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Salesmans.Name "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView3.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView3.Columns(0).HeaderText = "Код продавца"
        DataGridView3.Columns(0).Width = 50
        DataGridView3.Columns(1).HeaderText = "имя продавца"
        DataGridView3.Columns(1).Width = 150
        DataGridView3.Columns(2).HeaderText = "E-mail продавца"
        DataGridView3.Columns(2).Width = 200
        DataGridView3.Columns(3).HeaderText = "Город продавца"
        DataGridView3.Columns(3).Width = 150
        DataGridView3.Columns(4).HeaderText = "Ответ ствен ный за WEB в городе"
        DataGridView3.Columns(4).Width = 50
        DataGridView3.Columns(5).HeaderText = "Дежур ный"
        DataGridView3.Columns(5).Width = 50
        DataGridView3.Columns(6).HeaderText = "Актив ный"
        DataGridView3.Columns(6).Width = 50
        DataGridView3.Columns(7).HeaderText = "Резервное поле 1"
        DataGridView3.Columns(7).Width = 100
        DataGridView3.Columns(8).HeaderText = "Резервное поле 2"
        DataGridView3.Columns(8).Width = 100
        DataGridView3.Columns(9).HeaderText = "Статус БД"
        DataGridView3.Columns(9).Width = 50
        DataGridView3.Columns(10).HeaderText = "Статус WEB"
        DataGridView3.Columns(10).Width = 50
        DataGridView3.Columns(11).HeaderText = "Статус Scala"
        DataGridView3.Columns(11).Width = 50

        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductGroup()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка Групп продуктов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT Code, Name, WEBName, RMStatus, WEBStatus "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "ORDER BY Code "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView4.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView4.Columns(0).HeaderText = "Код группы"
        DataGridView4.Columns(0).Width = 100
        DataGridView4.Columns(1).HeaderText = "Название из Scala"
        DataGridView4.Columns(1).Width = 420
        DataGridView4.Columns(2).HeaderText = "название для WEB"
        DataGridView4.Columns(2).Width = 420
        DataGridView4.Columns(3).HeaderText = "Статус Scala"
        DataGridView4.Columns(3).Width = 50
        DataGridView4.Columns(4).HeaderText = "Статус WEB"
        DataGridView4.Columns(4).Width = 50

        DataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroup()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно подгрупп продуктов списка групп продуктов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---тут выводим группы, а сами подгруппы - по изменению выбора группы
        MySQLStr = "SELECT Code, Name, WEBName, RMStatus, WEBStatus "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "ORDER BY Code "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView5.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView5.Columns(0).HeaderText = "Код группы"
        DataGridView5.Columns(0).Width = 100
        DataGridView5.Columns(1).HeaderText = "Название из Scala"
        DataGridView5.Columns(1).Width = 420
        DataGridView5.Columns(2).HeaderText = "название для WEB"
        DataGridView5.Columns(2).Width = 420
        DataGridView5.Columns(3).HeaderText = "Статус Scala"
        DataGridView5.Columns(3).Width = 50
        DataGridView5.Columns(4).HeaderText = "Статус WEB"
        DataGridView5.Columns(4).Width = 50

        DataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroupDetail()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка подрупп продуктов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If DataGridView5.SelectedRows.Count > 0 Then
            MySQLStr = "SELECT SubgroupID, SubgroupCode, GroupCode, Name, Description, Rezerv1, RMStatus, WEBStatus "
            MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value & "') "
            MySQLStr = MySQLStr & "ORDER BY SubgroupCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView6.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView6.Columns(0).HeaderText = "ID под группы"
            DataGridView6.Columns(0).Width = 70
            DataGridView6.Columns(1).HeaderText = "Код под группы"
            DataGridView6.Columns(1).Width = 70
            DataGridView6.Columns(2).HeaderText = "Код группы"
            DataGridView6.Columns(2).Width = 70
            DataGridView6.Columns(3).HeaderText = "Название подгруппы"
            DataGridView6.Columns(3).Width = 320
            DataGridView6.Columns(4).HeaderText = "Описание подгруппы"
            DataGridView6.Columns(4).Width = 320
            DataGridView6.Columns(5).HeaderText = "Резервное поле"
            DataGridView6.Columns(5).Width = 100
            DataGridView6.Columns(6).HeaderText = "Статус Scala"
            DataGridView6.Columns(6).Width = 50
            DataGridView6.Columns(7).HeaderText = "Статус WEB"
            DataGridView6.Columns(7).Width = 50
        End If
        DataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProducts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно товаров списка групп продуктов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---тут выводим группы, а сами подгруппы - по изменению выбора группы
        MySQLStr = "SELECT Code, Name "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemGroup "
        MySQLStr = MySQLStr & "ORDER BY Code "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView7.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView7.Columns(0).HeaderText = "Код группы"
        DataGridView7.Columns(0).Width = 100
        DataGridView7.Columns(1).HeaderText = "Название из Scala"
        DataGridView7.Columns(1).Width = 390

        DataGridView7.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно товаров списка подгрупп продуктов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If DataGridView7.SelectedRows.Count > 0 Then
            MySQLStr = "SELECT SubgroupID, SubgroupCode, Name "
            MySQLStr = MySQLStr & "FROM  tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "WHERE (GroupCode = N'" & Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value & "') "
            MySQLStr = MySQLStr & "ORDER BY SubgroupCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView8.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView8.Columns(0).HeaderText = "ID подгруппы"
            DataGridView8.Columns(0).Width = 0
            DataGridView8.Columns(0).Visible = False
            DataGridView8.Columns(1).HeaderText = "Код подгруппы"
            DataGridView8.Columns(1).Width = 100
            DataGridView8.Columns(2).HeaderText = "Название подгруппы"
            DataGridView8.Columns(2).Width = 390

        End If
        DataGridView8.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProduct_IN_Subgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно товаров списка товаров, принадлежащих к выбранной подгруппе товаров
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If DataGridView7.SelectedRows.Count > 0 And DataGridView8.SelectedRows.Count > 0 Then


            'MySQLStr = "SELECT     tbl_WEB_Pictures.PictureSmall, tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, "
            'MySQLStr = MySQLStr & "tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.Country, tbl_WEB_Items.WEBName, tbl_WEB_Items.Description, "
            'MySQLStr = MySQLStr & "tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, tbl_WEB_Items.RMStatus, tbl_WEB_Items.WEBStatus "
            'MySQLStr = MySQLStr & "FROM         tbl_WEB_Items INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.ManufacturerItemCode = tbl_WEB_Pictures.PictureName "
            'MySQLStr = MySQLStr & "WHERE     (tbl_WEB_Items.GroupCode = N'" & Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value & "') AND (tbl_WEB_Items.SubGroupCode = N'" & Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value & "') "
            'MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

            MySQLStr = "SELECT tbl_WEB_Pictures.PictureSmall, tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.Country, tbl_WEB_Items.WEBName, tbl_WEB_Items.Description, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, tbl_WEB_Items.RMStatus, tbl_WEB_Items.WEBStatus "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value & "') "
            MySQLStr = MySQLStr & "AND (tbl_WEB_Items.SubGroupCode = N'" & Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

            DataGridView10.RowTemplate.MinimumHeight = 35

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView10.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView10.Columns(0).HeaderText = "Карт инка"
            DataGridView10.Columns(0).Width = 35
            DataGridView10.Columns(1).HeaderText = "Код запаса в Scala"
            DataGridView10.Columns(1).Width = 100
            DataGridView10.Columns(2).HeaderText = "Имя запаса в Scala"
            DataGridView10.Columns(2).Width = 300
            DataGridView10.Columns(3).HeaderText = "Производитель"
            DataGridView10.Columns(3).Width = 150
            DataGridView10.Columns(4).HeaderText = "Код запаса производителя"
            DataGridView10.Columns(4).Width = 150
            DataGridView10.Columns(5).HeaderText = "Страна"
            DataGridView10.Columns(5).Width = 100
            DataGridView10.Columns(6).HeaderText = "Имя запаса для WEB"
            DataGridView10.Columns(6).Width = 300
            DataGridView10.Columns(7).HeaderText = "Описание запаса"
            DataGridView10.Columns(7).Width = 300
            DataGridView10.Columns(8).HeaderText = "Складской ассортимент"
            DataGridView10.Columns(8).Width = 70
            DataGridView10.Columns(9).HeaderText = "Единица измерения"
            DataGridView10.Columns(9).Width = 70
            DataGridView10.Columns(10).HeaderText = "резервное поле2"
            DataGridView10.Columns(10).Width = 100
            DataGridView10.Columns(11).HeaderText = "Статус Scala"
            DataGridView10.Columns(11).Width = 40
            DataGridView10.Columns(12).HeaderText = "Статус WEB"
            DataGridView10.Columns(12).Width = 40

        End If
        DataGridView10.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView10.MultiSelect = True
    End Sub

    Private Sub LoadProduct_NO_Subgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно товаров списка товаров, не принадлежащих ни к одной подгруппе товаров
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If DataGridView7.SelectedRows.Count > 0 Then

            'MySQLStr = "SELECT     tbl_WEB_Pictures.PictureSmall, tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, "
            'MySQLStr = MySQLStr & "tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.Country, tbl_WEB_Items.WEBName, tbl_WEB_Items.Description, "
            'MySQLStr = MySQLStr & "tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, tbl_WEB_Items.RMStatus, tbl_WEB_Items.WEBStatus "
            'MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.ManufacturerItemCode = tbl_WEB_Pictures.PictureName "
            'MySQLStr = MySQLStr & "WHERE     (tbl_WEB_Items.GroupCode = N'" & Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value & "') AND (tbl_WEB_Items.SubGroupCode = N'') "
            'MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

            MySQLStr = "SELECT tbl_WEB_Pictures.PictureSmall, tbl_WEB_Items.Code, tbl_WEB_Items.Name, tbl_WEB_Manufacturers.Name AS ManufacturerName, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.ManufacturerItemCode, tbl_WEB_Items.Country, tbl_WEB_Items.WEBName, tbl_WEB_Items.Description, "
            MySQLStr = MySQLStr & "tbl_WEB_Items.WHAssortiment, tbl_WEB_Items.UOM, tbl_WEB_Items.Rezerv, tbl_WEB_Items.RMStatus, tbl_WEB_Items.WEBStatus "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Items INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures ON tbl_WEB_Items.Code = tbl_WEB_Pictures.ScalaItemCode "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.GroupCode = N'" & Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value & "') AND (tbl_WEB_Items.SubGroupCode = N'') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_Items.Code "

            DataGridView9.RowTemplate.MinimumHeight = 35

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView9.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView9.Columns(0).HeaderText = "Карт инка"
            DataGridView9.Columns(0).Width = 40
            DataGridView9.Columns(1).HeaderText = "Код запаса в Scala"
            DataGridView9.Columns(1).Width = 100
            DataGridView9.Columns(2).HeaderText = "Имя запаса в Scala"
            DataGridView9.Columns(2).Width = 300
            DataGridView9.Columns(3).HeaderText = "Производитель"
            DataGridView9.Columns(3).Width = 150
            DataGridView9.Columns(4).HeaderText = "Код запаса производителя"
            DataGridView9.Columns(4).Width = 150
            DataGridView9.Columns(5).HeaderText = "Страна"
            DataGridView9.Columns(5).Width = 100
            DataGridView9.Columns(6).HeaderText = "Имя запаса для WEB"
            DataGridView9.Columns(6).Width = 300
            DataGridView9.Columns(7).HeaderText = "Описание запаса"
            DataGridView9.Columns(7).Width = 300
            DataGridView9.Columns(8).HeaderText = "Складской ассортимент"
            DataGridView9.Columns(8).Width = 70
            DataGridView9.Columns(9).HeaderText = "Единица измерения"
            DataGridView9.Columns(9).Width = 70
            DataGridView9.Columns(10).HeaderText = "резервное поле"
            DataGridView9.Columns(10).Width = 100
            DataGridView9.Columns(11).HeaderText = "Статус Scala"
            DataGridView9.Columns(11).Width = 40
            DataGridView9.Columns(12).HeaderText = "Статус WEB"
            DataGridView9.Columns(12).Width = 40

        End If
        DataGridView9.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView9.MultiSelect = True
    End Sub

    Private Sub LoadCustomers()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка Клиентов
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        MySQLStr = "SELECT Code, Name, Address, Discount, Case WHEN WorkOverWEB = 1 THEN 'Да' ELSE '' END as WorkOverWEB, BasePrice "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
        MySQLStr = MySQLStr & "ORDER BY Code "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView11.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView11.Columns(0).HeaderText = "Код клиента"
        DataGridView11.Columns(0).Width = 100
        DataGridView11.Columns(1).HeaderText = "Название клиента"
        DataGridView11.Columns(1).Width = 250
        DataGridView11.Columns(2).HeaderText = "Адрес"
        DataGridView11.Columns(2).Width = 520
        DataGridView11.Columns(3).HeaderText = "Общая скидка"
        DataGridView11.Columns(3).Width = 60
        DataGridView11.Columns(4).HeaderText = "Работает через WEB"
        DataGridView11.Columns(4).Width = 60
        DataGridView11.Columns(5).HeaderText = "Базовый прайс"
        DataGridView11.Columns(5).Width = 60

        DataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadDiscountsHeader()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка в окно списка клиентов, работающих через WEB
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка подгрупп
        Dim MyDs As New DataSet

        '---------------Список клиентов
        MySQLStr = "SELECT Code, LTRIM(RTRIM(LTRIM(RTRIM(Code)) + ' ' + LTRIM(RTRIM(Name)))) AS Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
        MySQLStr = MySQLStr & "WHERE (WorkOverWEB = 1) "
        MySQLStr = MySQLStr & "ORDER BY Name "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "Code"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        If ComboBox1.SelectedValue = Nothing Then
        Else
            MySQLStr = "SELECT Discount "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & ComboBox1.SelectedValue & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                TextBox3.Text = ""
            Else
                TextBox3.Text = Declarations.MyRec.Fields("Discount").Value
                trycloseMyRec()
            End If
        End If
    End Sub

    Private Sub CheckCitiesButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для городов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button20.Enabled = False
            Button21.Enabled = False
        Else
            Button20.Enabled = True
            Button21.Enabled = True
        End If
    End Sub

    Private Sub CheckManufacturersButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для производителей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button23.Enabled = False
            Button24.Enabled = False
            Button25.Enabled = False
        Else
            If DataGridView2.SelectedRows.Item(0).Cells(4).Value = 2 Or DataGridView2.SelectedRows.Item(0).Cells(5).Value = 2 Then
                Button23.Enabled = False
            Else
                Button23.Enabled = True
            End If
            Button24.Enabled = True
            Button25.Enabled = True
        End If
    End Sub

    Private Sub CheckSalesmansButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для продавцов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView3.SelectedRows.Count = 0 Then
            Button27.Enabled = False
            Button28.Enabled = False
            Button29.Enabled = False
        Else
            If DataGridView3.SelectedRows.Item(0).Cells(11).Value = 2 Then
                Button27.Enabled = False
            Else
                Button27.Enabled = True
            End If
            Button28.Enabled = True
            Button29.Enabled = True
        End If
    End Sub

    Private Sub CheckProductGroupButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для групп продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView4.SelectedRows.Count = 0 Then
            Button31.Enabled = False
            Button32.Enabled = False
            Button33.Enabled = False
        Else
            If DataGridView4.SelectedRows.Item(0).Cells(3).Value = 2 Or DataGridView4.SelectedRows.Item(0).Cells(4).Value = 2 Then
                Button33.Enabled = False
            Else
                Button33.Enabled = True
            End If
            Button31.Enabled = True
            Button32.Enabled = True
        End If
    End Sub

    Private Sub CheckProductSubGroupButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для подгрупп продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView6.SelectedRows.Count = 0 Then
            Button35.Enabled = False
            Button36.Enabled = False
            Button37.Enabled = False
        Else
            If DataGridView6.SelectedRows.Item(0).Cells(6).Value = 2 Or DataGridView6.SelectedRows.Item(0).Cells(7).Value = 2 Then
                Button35.Enabled = False
            Else
                Button35.Enabled = True
            End If
            Button36.Enabled = True
            Button37.Enabled = True
        End If
    End Sub

    Private Sub CheckProductButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '----------кнопки перемещения в/из подгруппы
        If DataGridView10.SelectedRows.Count = 0 Then
            Button42.Enabled = False
        Else
            Button42.Enabled = True
        End If

        If DataGridView9.SelectedRows.Count = 0 Or DataGridView8.SelectedRows.Count = 0 Then
            Button41.Enabled = False
        Else
            Button41.Enabled = True
        End If

        '---------Кнопки редактирования запасов
        If DataGridView10.SelectedRows.Count = 0 Then
            Button44.Enabled = False
        Else
            Button44.Enabled = True
        End If

        If DataGridView9.SelectedRows.Count = 0 Then
            Button43.Enabled = False
        Else
            Button43.Enabled = True
        End If

        '----------Кнопки выгрузки в Excel
        If DataGridView7.SelectedRows.Count = 0 Then
            Button45.Enabled = False
        Else
            Button45.Enabled = True
        End If

        If DataGridView8.SelectedRows.Count = 0 Or DataGridView10.SelectedRows.Count = 0 Then
            Button46.Enabled = False
        Else
            Button46.Enabled = True
        End If

        If DataGridView7.SelectedRows.Count = 0 Or DataGridView9.SelectedRows.Count = 0 Then
            Button47.Enabled = False
        Else
            Button47.Enabled = True
        End If
    End Sub

    Private Sub CheckCustomerButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для городов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView11.SelectedRows.Count = 0 Then
            Button52.Enabled = False
        Else
            Button52.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отрисовка подписей к Tab горизонтально
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выделенного города
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim RemFlag As Integer          'флаг - можно ли удалять запись

        RemFlag = 0
        MySQLStr = "SELECT COUNT(Code) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Salesmans "
        MySQLStr = MySQLStr & "WHERE(City = " & Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            RemFlag = 1
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                trycloseMyRec()
            Else
                trycloseMyRec()
                RemFlag = 1
            End If
        End If
        If RemFlag = 0 Then
            MsgBox("Текущую запись нельзя удалить, так как есть записи по продавцам, ссылающиеся на данный город.", MsgBoxStyle.Critical, "Внимание!")
        Else
            MySQLStr = "DELETE FROM tbl_WEB_Cities "
            MySQLStr = MySQLStr & "WHERE (ID = " & Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---загрузка данных
            LoadCities()
            CheckCitiesButtons()
        End If

    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание города
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCity = New City
        MyCity.StartParam = "Create"
        Declarations.MyCityID = 0
        MyCity.ShowDialog()
        '---загрузка данных
        LoadCities()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyCityID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
            End If
        Next
        CheckCitiesButtons()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование города
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCity = New City
        MyCity.StartParam = "Edit"
        Declarations.MyCityID = Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyCity.ShowDialog()
        '---загрузка данных
        LoadCities()
        '---текущей строкой сделать измененную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyCityID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
            End If
        Next
        CheckCitiesButtons()
    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк производителей в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView2.Rows(e.RowIndex)
        If row.Cells(4).Value = 1 Or row.Cells(4).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(4).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(5).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по поставщикам из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Manufacturers_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("Загрузка информации по производителям из Scala произведена", MsgBoxStyle.Information, "Внимание!")
        LoadManufacturers()
        CheckManufacturersButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyManufacturer = New Manufacturer
        Declarations.MyManufacturerID = Me.DataGridView2.SelectedRows.Item(0).Cells(0).Value
        MyManufacturer.ShowDialog()
        '---загрузка данных
        LoadManufacturers()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Item(0, i).Value = Declarations.MyManufacturerID Then
                DataGridView2.CurrentCell = DataGridView2.Item(1, i)
            End If
        Next
        CheckManufacturersButtons()
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора производителя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckManufacturersButtons()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора города
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckCitiesButtons()
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка производителей в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadManufacturersToLO()
        Else
            UploadManufacturersToExcel()
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка производителей в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadManufacturersToLO()
        Else
            UploadManufacturersToExcel()
        End If
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по производителям из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadManufacturersFromLO()
        Else
            LoadManufacturersFromExcel()
        End If
        MsgBox("Загрузка данных по производителям из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadManufacturers()
        CheckManufacturersButtons()
    End Sub

    Private Sub DataGridView3_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк продавцов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView3.Rows(e.RowIndex)
        If row.Cells(11).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(9).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(6).Value = "" Then
            row.DefaultCellStyle.BackColor = Color.LightBlue
        ElseIf row.Cells(9).Value = 1 Or row.Cells(9).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(10).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub DataGridView3_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора продавца
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSalesmansButtons()
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по продавцам из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Salesmans_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("Загрузка информации по продавцам из Scala произведена", MsgBoxStyle.Information, "Внимание!")
        LoadSalesmans()
        CheckSalesmansButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование продавца
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySalesman = New Salesman
        Declarations.MySalesmanID = Me.DataGridView3.SelectedRows.Item(0).Cells(0).Value
        MySalesman.ShowDialog()
        '---загрузка данных
        LoadSalesmans()
        '---текущей строкой сделать измененную
        For i As Integer = 0 To DataGridView3.Rows.Count - 1
            If DataGridView3.Item(0, i).Value = Declarations.MySalesmanID Then
                DataGridView3.CurrentCell = DataGridView3.Item(1, i)
            End If
        Next
        CheckSalesmansButtons()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продавцов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadSalesmansToLO()
        Else
            UploadSalesmansToExcel()
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продавцов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadSalesmansToLO()
        Else
            UploadSalesmansToExcel()
        End If
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по продавцам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadSalesmansFromLO()
        Else
            LoadSalesmansFromExcel()
        End If
        MsgBox("Загрузка данных по продавцам из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadSalesmans()
        CheckSalesmansButtons()
    End Sub

    Private Sub DataGridView4_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк групп продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView4.Rows(e.RowIndex)
        If row.Cells(3).Value = 1 Or row.Cells(3).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(3).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(4).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView4_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView4.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора группы продукта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProductGroupButtons()
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по группам продуктов из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_ItemGroups_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("Загрузка информации по группам продуктов из Scala произведена", MsgBoxStyle.Information, "Внимание!")
        LoadProductGroup()
        CheckProductGroupButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование группы продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductGroup = New ProductGroup
        Declarations.MyProductGroupID = Me.DataGridView4.SelectedRows.Item(0).Cells(0).Value
        MyProductGroup.ShowDialog()
        '---загрузка данных
        LoadProductGroup()
        '---текущей строкой сделать измененную
        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Item(0, i).Value = Declarations.MyProductGroupID Then
                DataGridView4.CurrentCell = DataGridView4.Item(1, i)
            End If
        Next
        CheckProductGroupButtons()
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка групп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductGroupsToLO()
        Else
            UploadProductGroupsToExcel()
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка групп продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductGroupsToLO()
        Else
            UploadProductGroupsToExcel()
        End If
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по группам товаров из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductGroupsFromLO()
        Else
            LoadProductGroupsFromExcel()
        End If
        MsgBox("Загрузка данных по группам продуктов из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadProductGroup()
        CheckProductGroupButtons()
    End Sub

    Private Sub DataGridView5_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк групп продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView5.Rows(e.RowIndex)
        If row.Cells(3).Value = 1 Or row.Cells(3).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(3).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(4).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView5_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView5.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора группы продукта в закладке подгрупп
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProductSubgroupDetail()
    End Sub

    Private Sub DataGridView6_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView6.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк подгрупп продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView6.Rows(e.RowIndex)
        If row.Cells(6).Value = 1 Or row.Cells(6).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(6).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(7).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView6_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView6.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора подгруппы продукта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProductSubGroupButtons()
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка подгрупп продуктов в Excel (только для текущей группы)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductSubGroupsToLO(Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value), Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(1).Value))
        Else
            UploadProductSubGroupsToExcel(Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value), Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(1).Value))
        End If

    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка подгрупп продуктов в Excel (для всех групп)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductSubGroupsToLO("", "")
        Else
            UploadProductSubGroupsToExcel("", "")
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка подгрупп продуктов в Excel (для всех групп)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductSubGroupsToLO("", "")
        Else
            UploadProductSubGroupsToExcel("", "")
        End If
    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по подгруппам товаров из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductSubGroupsFromLO()
        Else
            LoadProductSubGroupsFromExcel()
        End If
        MsgBox("Загрузка данных по подгруппам продуктов из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadProductSubgroup()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание подгруппы продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductSubGroup = New ProductSubGroup
        MyProductSubGroup.StartParam = "Create"
        Declarations.MyProductGroupID = Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductSubGroupID = "N"
        MyProductSubGroup.ShowDialog()
        '---загрузка данных
        LoadProductSubgroupDetail()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView6.Rows.Count - 1
            If Trim(DataGridView6.Item(1, i).Value.ToString) = Declarations.MyProductSubGroupID Then
                DataGridView6.CurrentCell = DataGridView6.Item(1, i)
            End If
        Next
        CheckProductSubGroupButtons()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование подгруппы продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductSubGroup = New ProductSubGroup
        MyProductSubGroup.StartParam = "Edit"
        Declarations.MyProductGroupID = Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductSubGroupID = Trim(Me.DataGridView6.SelectedRows.Item(0).Cells(1).Value)
        MyProductSubGroup.ShowDialog()
        '---загрузка данных
        LoadProductSubgroupDetail()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView6.Rows.Count - 1
            If Trim(DataGridView6.Item(1, i).Value.ToString) = Declarations.MyProductSubGroupID Then
                DataGridView6.CurrentCell = DataGridView6.Item(1, i)
            End If
        Next
        CheckProductSubGroupButtons()
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выделенной подгруппы продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim RemFlag As Integer          'флаг - можно ли удалять запись

        RemFlag = 0
        MySQLStr = "SELECT COUNT(Code) AS CC "
        MySQLStr = MySQLStr & "FROM  tbl_WEB_Items "
        MySQLStr = MySQLStr & "WHERE(GroupCode = N'" & Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value & "') "
        MySQLStr = MySQLStr & "AND (SubGroupCode = N'" & Me.DataGridView6.SelectedRows.Item(0).Cells(1).Value & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            RemFlag = 1
            trycloseMyRec()
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                trycloseMyRec()
            Else
                trycloseMyRec()
                RemFlag = 1
            End If
        End If
        If RemFlag = 0 Then
            MsgBox("Текущую подгруппу нельзя удалить, так как есть товары, входящие в данную подгруппу.", MsgBoxStyle.Critical, "Внимание!")
        Else
            'MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            'MySQLStr = MySQLStr & "WHERE(GroupCode = N'" & Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value & "') "
            'MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & Me.DataGridView6.SelectedRows.Item(0).Cells(1).Value & "') "
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET RMStatus = 2, WEBStatus = 2 "
            MySQLStr = MySQLStr & "WHERE(GroupCode = N'" & Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value & "') "
            MySQLStr = MySQLStr & "AND (SubgroupCode = N'" & Me.DataGridView6.SelectedRows.Item(0).Cells(1).Value & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---загрузка данных
            LoadProductSubgroupDetail()
            CheckProductSubGroupButtons()
        End If
    End Sub

    Private Sub DataGridView7_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView7.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора группы продукта в закладке товара
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProductSubgroup_IN_Items()
        If DataGridView8.SelectedRows.Count = 0 Then
            DataGridView10.DataSource = ""
        End If
        LoadProduct_NO_Subgroup_IN_Items()
        CheckProductButtons()
    End Sub

    Private Sub DataGridView8_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView8.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора подгруппы продукта в закладке товара
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProduct_IN_Subgroup_IN_Items()
        CheckProductButtons()
    End Sub

    Private Sub DataGridView9_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView9.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView9.Rows(e.RowIndex)
        If row.Cells(11).Value = 1 Or row.Cells(11).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(11).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(12).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView10_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView10.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк продуктов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView10.Rows(e.RowIndex)
        If row.Cells(11).Value = 1 Or row.Cells(11).Value = 3 Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        ElseIf row.Cells(11).Value = 2 Then
            row.DefaultCellStyle.BackColor = Color.LightGray
        ElseIf row.Cells(12).Value <> 0 Then
            row.DefaultCellStyle.BackColor = Color.LightPink
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по группам продуктов из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Items_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("Загрузка информации по продуктам из Scala произведена", MsgBoxStyle.Information, "Внимание!")
        LoadProducts()
        CheckProductButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление запасов из выделенной подгруппы
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim i As Integer
        Dim MySQLStr As String

        If DataGridView10.SelectedRows.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            For i = 0 To DataGridView10.SelectedRows.Count - 1
                '--------Обновление информации в запасе
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET SubGroupCode = N'', "
                MySQLStr = MySQLStr & "RMStatus = 2, "
                MySQLStr = MySQLStr & "WEBStatus = 2 "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView10.SelectedRows.Item(i).Cells(1).Value & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '------Обновление информации в истории
                MySQLStr = "DELETE FROM tbl_WEB_Items_InSubGroupHistory "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView10.SelectedRows.Item(i).Cells(1).Value & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                'MsgBox(DataGridView10.SelectedRows.Item(i).Cells(0).Value, MsgBoxStyle.Information, "Attention!")
            Next i
            LoadProduct_IN_Subgroup_IN_Items()
            LoadProduct_NO_Subgroup_IN_Items()
            CheckProductButtons()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление запасов в выделенную подгруппу
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer
        Dim MySQLStr As String

        If DataGridView9.SelectedRows.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            For i = 0 To DataGridView9.SelectedRows.Count - 1
                '------Обновление информации в запасе
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET SubGroupCode = N'" & Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value & "', "
                MySQLStr = MySQLStr & "RMStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE 1 END, "
                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE 1 END "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView9.SelectedRows.Item(i).Cells(1).Value & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '------Обновление информации в истории
                MySQLStr = "DELETE FROM tbl_WEB_Items_InSubGroupHistory "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView9.SelectedRows.Item(i).Cells(1).Value & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                MySQLStr = "INSERT INTO tbl_WEB_Items_InSubGroupHistory "
                MySQLStr = MySQLStr & "(Code, SubGroupCode) "
                MySQLStr = MySQLStr & "VALUES (N'" & DataGridView9.SelectedRows.Item(i).Cells(1).Value & "', "
                MySQLStr = MySQLStr & "N'" & Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)


                'MsgBox(DataGridView9.SelectedRows.Item(i).Cells(0).Value, MsgBoxStyle.Information, "Attention!")
            Next i
            LoadProduct_IN_Subgroup_IN_Items()
            LoadProduct_NO_Subgroup_IN_Items()
            CheckProductButtons()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Все запасы только для текущей группы)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            "", _
            "", _
            2)
        Else
            UploadProductsToExcel(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            "", _
            "", _
            2)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Все запасы только для текущей группы, для текущей подгруппы)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            Trim(Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value), _
            Trim(Me.DataGridView8.SelectedRows.Item(0).Cells(2).Value), _
            1)
        Else
            UploadProductsToExcel(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            Trim(Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value), _
            Trim(Me.DataGridView8.SelectedRows.Item(0).Cells(2).Value), _
            1)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button47.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Для текущей группы, не входящие ни в одну подгруппу)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            "", _
            "", _
            0)
        Else
            UploadProductsToExcel(Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value), _
            Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(1).Value), _
            "", _
            "", _
            0)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button48.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Все запасы)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO("", _
            "", _
            "", _
            "", _
            2)
        Else
            UploadProductsToExcel("", _
            "", _
            "", _
            "", _
            2)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button49.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Все запасы без подгрупп)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO("", "", "", "", 0)
        Else
            UploadProductsToExcel("", "", "", "", 0)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка продуктов в Excel (Все запасы) с закладки операции
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO("", "", "", "", 2)
        Else
            UploadProductsToExcel("", "", "", "", 2)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button50.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по товарам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductsFromLO()
        Else
            LoadProductsFromExcel()
        End If

        MsgBox("Загрузка данных по продуктам из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        CheckProductButtons()
        Cursor = Cursors.Default
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProduct = New Product
        MyProduct.StartParam = "NotInGroup"
        MyProduct.MyGroup = Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductID = Trim(Me.DataGridView9.SelectedRows.Item(0).Cells(1).Value)
        MyProduct.ShowDialog()
        '---загрузка данных
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView9.Rows.Count - 1
            If Trim(DataGridView9.Item(0, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView9.CurrentCell = DataGridView9.Item(0, i)
            End If
        Next
        CheckProductSubGroupButtons()
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProduct = New Product
        MyProduct.StartParam = "InGroup"
        MyProduct.MyGroup = Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductID = Trim(Me.DataGridView10.SelectedRows.Item(0).Cells(1).Value)
        MyProduct.ShowDialog()
        '---загрузка данных
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView10.Rows.Count - 1
            If Trim(DataGridView10.Item(0, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView10.CurrentCell = DataGridView10.Item(0, i)
            End If
        Next
        CheckProductSubGroupButtons()
    End Sub

    Private Sub DataGridView11_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView11.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк клиентов в зависимости от статуса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView11.Rows(e.RowIndex)
        If row.Cells(4).Value = "Да" Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView11_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView11.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора клиента в закладке клиенты
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckCustomerButtons()
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию покупателей в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox1.Select()
        Else
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button51.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по Клиентам из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Clients_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("Загрузка информации по клиентам из Scala произведена", MsgBoxStyle.Information, "Внимание!")
        LoadCustomers()
        CheckCustomerButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button52.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomer = New Customer
        Declarations.MyCustomerID = Trim(Me.DataGridView11.SelectedRows.Item(0).Cells(0).Value)
        MyCustomer.ShowDialog()
        '---загрузка данных
        LoadCustomers()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView11.Rows.Count - 1
            If Trim(DataGridView11.Item(0, i).Value.ToString) = Declarations.MyCustomerID Then
                DataGridView11.CurrentCell = DataGridView11.Item(0, i)
            End If
        Next
        CheckCustomerButtons()
    End Sub

    Private Sub TabControl2_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl2.Selecting
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор закладки
        '// для выбранной закладки выводим данные
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        Select Case sender.selectedtab.text
            Case "Скидки на группы продуктов"
                LoadGroupDiscount()
                CheckGroupDiscountsButtons()
            Case "Скидки на подгруппы продуктов"
                LoadSubgroupDiscounts()
                CheckSubgroupDiscountsButtons()
            Case "Скидки на продукты"
                LoadItemDiscounts()
                CheckItemDiscountsButtons()
            Case "Согласованный ассортимент"
                LoadAgreedRange()
                CheckAgreedRangeButtons()
            Case Else
        End Select
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора клиента, для которого вводятся скидки
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If ComboBox1.SelectedValue = Nothing Then
            TextBox3.Text = ""
        Else
            MySQLStr = "SELECT tbl_WEB_Clients.Discount, ISNULL(View_1.Name,'') AS BasePriceInfo "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Clients LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT DISTINCT SY240300.SY24002 AS Code, SY240300.SY24002 + ' ' + SY240300.SY24003 AS Name "
            MySQLStr = MySQLStr & "FROM SY240300 INNER JOIN "
            MySQLStr = MySQLStr & "SC390300 ON SY240300.SY24002 = SC390300.SC39002 "
            MySQLStr = MySQLStr & "WHERE (SY240300.SY24001 = N'IL')) AS View_1 ON tbl_WEB_Clients.BasePrice = View_1.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Clients.Code = N'" & ComboBox1.SelectedValue & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                TextBox3.Text = ""
                TextBox4.Text = ""
            Else
                TextBox3.Text = Declarations.MyRec.Fields("Discount").Value
                TextBox4.Text = Declarations.MyRec.Fields("BasePriceInfo").Value
                trycloseMyRec()
            End If
        End If

        If TabControl2.SelectedTab.Text = "Скидки на группы продуктов" Then
            LoadGroupDiscount()
            CheckGroupDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "Скидки на подгруппы продуктов" Then
            LoadSubgroupDiscounts()
            CheckSubgroupDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "Скидки на продукты" Then
            LoadItemDiscounts()
            CheckItemDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "Согласованный ассортимент" Then
            LoadAgreedRange()
            CheckAgreedRangeButtons()
        Else

        End If
    End Sub

    Private Sub LoadGroupDiscount()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка скидок по группе
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox1.SelectedValue = Nothing Then
            DataGridView12.DataSource = ""
        Else
            MySQLStr = "SELECT tbl_WEB_DiscountGroup.ID, tbl_WEB_DiscountGroup.GroupCode, ISNULL(tbl_WEB_ItemGroup.Name, N'') AS Name, tbl_WEB_DiscountGroup.Discount "
            MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountGroup LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountGroup.GroupCode = tbl_WEB_ItemGroup.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountGroup.ClientCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountGroup.GroupCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView12.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView12.Columns(0).HeaderText = "ID группы товаров"
            DataGridView12.Columns(0).Width = 0
            DataGridView12.Columns(0).Visible = False
            DataGridView12.Columns(1).HeaderText = "Код группы товаров"
            DataGridView12.Columns(1).Width = 150
            DataGridView12.Columns(2).HeaderText = "Название группы товаров"
            DataGridView12.Columns(2).Width = 800
            DataGridView12.Columns(3).HeaderText = "Скидка (%)"
            DataGridView12.Columns(3).Width = 100
        End If

        DataGridView12.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadSubgroupDiscounts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка скидок по подгруппе
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox1.SelectedValue = Nothing Then
            DataGridView12.DataSource = ""
        Else
            MySQLStr = "SELECT  tbl_WEB_DiscountSubgroup.ID, tbl_WEB_DiscountSubgroup.GroupCode, tbl_WEB_ItemGroup.Name, tbl_WEB_DiscountSubgroup.SubgroupCode, "
            MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup.Name AS Expr1, tbl_WEB_DiscountSubgroup.Discount "
            MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountSubgroup LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
            MySQLStr = MySQLStr & "tbl_WEB_DiscountSubgroup.SubgroupCode = tbl_WEB_ItemSubGroup.SubgroupCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemGroup ON tbl_WEB_DiscountSubgroup.GroupCode = tbl_WEB_ItemGroup.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountSubgroup.ClientCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountSubgroup.GroupCode, tbl_WEB_DiscountSubgroup.SubgroupCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView13.DataSource = MyDs.Tables(0)
                '---фильтр
                MyBS.DataSource = MyDs
                MyBS.DataMember = MyDs.Tables(0).TableName
                DataGridView13.DataSource = MyBS
                '---конец фильтра

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView13.Columns(0).HeaderText = "ID группы товаров"
            DataGridView13.Columns(0).Width = 0
            DataGridView13.Columns(0).Visible = False
            DataGridView13.Columns(1).HeaderText = "Код группы товаров"
            DataGridView13.Columns(1).Width = 80
            DataGridView13.Columns(2).HeaderText = "Название группы товаров"
            DataGridView13.Columns(2).Width = 400
            DataGridView13.Columns(3).HeaderText = "Код подгруппы товаров"
            DataGridView13.Columns(3).Width = 80
            DataGridView13.Columns(4).HeaderText = "Название подгруппы товаров"
            DataGridView13.Columns(4).Width = 400
            DataGridView13.Columns(5).HeaderText = "Скидка (%)"
            DataGridView13.Columns(5).Width = 90
        End If

        DataGridView13.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadItemDiscounts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка скидок по товарам
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox1.SelectedValue = Nothing Then
            DataGridView14.DataSource = ""
        Else
            MySQLStr = "SELECT tbl_WEB_DiscountItem.ID, tbl_WEB_DiscountItem.ItemCode, tbl_WEB_Items.Name, tbl_WEB_DiscountItem.Discount "
            MySQLStr = MySQLStr & "FROM tbl_WEB_DiscountItem LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_DiscountItem.ItemCode = tbl_WEB_Items.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_DiscountItem.ClientCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_DiscountItem.ItemCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView14.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView14.Columns(0).HeaderText = "ID записи"
            DataGridView14.Columns(0).Width = 0
            DataGridView14.Columns(0).Visible = False
            DataGridView14.Columns(1).HeaderText = "Код товара"
            DataGridView14.Columns(1).Width = 150
            DataGridView14.Columns(2).HeaderText = "Название товара"
            DataGridView14.Columns(2).Width = 800
            DataGridView14.Columns(3).HeaderText = "Скидка (%)"
            DataGridView14.Columns(3).Width = 100
        End If

        DataGridView14.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadAgreedRange()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в окно списка согласованного ассортимента (БЕЗ НДС)
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox1.SelectedValue = Nothing Then
            DataGridView12.DataSource = ""
        Else
            MySQLStr = "SELECT tbl_WEB_AgreedRange.ID, tbl_WEB_AgreedRange.ItemCode, ISNULL(tbl_WEB_Items.Name,'') AS Name, tbl_WEB_AgreedRange.AgreedPrice, ISNULL(SYCD0100.SYCD009,'') AS CurrName "
            MySQLStr = MySQLStr & "FROM tbl_WEB_AgreedRange LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SYCD0100 ON tbl_WEB_AgreedRange.CurrCode = SYCD0100.SYCD001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_Items ON tbl_WEB_AgreedRange.ItemCode = tbl_WEB_Items.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_AgreedRange.ClientCode = N'" & Trim(ComboBox1.SelectedValue) & "') "
            MySQLStr = MySQLStr & "ORDER BY tbl_WEB_AgreedRange.ItemCode "

            InitMyConn(False)
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView15.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView15.Columns(0).HeaderText = "ID записи"
            DataGridView15.Columns(0).Width = 0
            DataGridView15.Columns(0).Visible = False
            DataGridView15.Columns(1).HeaderText = "Код товара"
            DataGridView15.Columns(1).Width = 150
            DataGridView15.Columns(2).HeaderText = "Название товара"
            DataGridView15.Columns(2).Width = 740
            DataGridView15.Columns(3).HeaderText = "цена (без НДС)"
            DataGridView15.Columns(3).Width = 80
            DataGridView15.Columns(4).HeaderText = "Валюта"
            DataGridView15.Columns(4).Width = 80
        End If

        DataGridView15.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub DataGridView12_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView12.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора скидки по группе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckGroupDiscountsButtons()
    End Sub

    Private Sub CheckGroupDiscountsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для скидок по группе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView12.SelectedRows.Count = 0 Then
            Button53.Enabled = False
            Button55.Enabled = False
            Button56.Enabled = False
        Else
            Button53.Enabled = True
            Button55.Enabled = True
            Button56.Enabled = True
        End If
    End Sub

    Private Sub CheckSubgroupDiscountsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для скидок по подгруппе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView13.SelectedRows.Count = 0 Then
            Button62.Enabled = False
            Button61.Enabled = False
            Button60.Enabled = False
        Else
            Button62.Enabled = True
            Button61.Enabled = True
            Button60.Enabled = True
        End If
    End Sub

    Private Sub CheckItemDiscountsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для скидок по товару
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView14.SelectedRows.Count = 0 Then
            Button65.Enabled = False
            Button66.Enabled = False
            Button67.Enabled = False
        Else
            Button65.Enabled = True
            Button66.Enabled = True
            Button67.Enabled = True
        End If
    End Sub

    Private Sub CheckAgreedRangeButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и установка статуса кнопок для согласованного ассортимента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView15.SelectedRows.Count = 0 Then
            Button72.Enabled = False
            Button71.Enabled = False
            Button70.Enabled = False
        Else
            Button72.Enabled = True
            Button71.Enabled = True
            Button70.Enabled = True
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание скидки на группу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountGroup = New DiscountGroup
        MyDiscountGroup.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountGroup.ShowDialog()
        '---загрузка данных
        LoadGroupDiscount()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView12.Rows.Count - 1
            If Trim(DataGridView12.Item(1, i).Value.ToString) = Declarations.MyProductGroupID Then
                DataGridView12.CurrentCell = DataGridView12.Item(1, i)
            End If
        Next
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub Button53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button53.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование скидки на группу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountGroup = New DiscountGroup
        MyDiscountGroup.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductGroupID = Trim(Me.DataGridView12.SelectedRows.Item(0).Cells(1).Value)
        MyDiscountGroup.ShowDialog()
        '---загрузка данных
        LoadGroupDiscount()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView12.Rows.Count - 1
            If Trim(DataGridView12.Item(1, i).Value.ToString) = Declarations.MyProductGroupID Then
                DataGridView12.CurrentCell = DataGridView12.Item(1, i)
            End If
        Next
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub Button55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button55.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выбранной скидки на группу
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountGroup "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView12.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---загрузка данных
        LoadGroupDiscount()
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub Button56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button56.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel скидок по группам для текущего клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadGroupDiscountToLO(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        Else
            UploadGroupDiscountToExcel(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button57.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по скидкам по группам товаров из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadGroupDiscountsFromLO()
        Else
            LoadGroupDiscountsFromExcel()
        End If
        MsgBox("Загрузка данных по скидкам по группам продуктов из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadGroupDiscount()
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub DataGridView13_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView13.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню выставления фильтра
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Declarations.MyFilterColumn = e.ColumnIndex
            ContextMenuStrip1.Show(MousePosition.X, MousePosition.Y)
        End If
    End Sub

    Private Sub DataGridView13_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView13.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора скидки по группе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button63.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание скидки на подгруппу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountSubgroup = New DiscountSubgroup
        MyDiscountSubgroup.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountSubgroup.ShowDialog()
        '---загрузка данных
        LoadSubgroupDiscounts()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView13.Rows.Count - 1
            If Trim(DataGridView13.Item(1, i).Value.ToString) = Declarations.MyProductGroupID _
                And Trim(DataGridView13.Item(3, i).Value.ToString) = Declarations.MyProductSubGroupID Then
                DataGridView13.CurrentCell = DataGridView13.Item(1, i)
            End If
        Next
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button62.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение скидки на подгруппу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountSubgroup = New DiscountSubgroup
        MyDiscountSubgroup.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductGroupID = Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(1).Value)
        Declarations.MyProductSubGroupID = Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(3).Value)
        MyDiscountSubgroup.ShowDialog()
        '---загрузка данных
        LoadSubgroupDiscounts()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView13.Rows.Count - 1
            If Trim(DataGridView13.Item(1, i).Value.ToString) = Declarations.MyProductGroupID _
                And Trim(DataGridView13.Item(3, i).Value.ToString) = Declarations.MyProductSubGroupID Then
                DataGridView13.CurrentCell = DataGridView13.Item(1, i)
            End If
        Next
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button61.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выбранной скидки на подгруппу
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountSubgroup "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---загрузка данных
        LoadSubgroupDiscounts()
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button60.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel скидок по подгруппам для текущего клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadSubgroupDiscountToLO(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        Else
            UploadSubgroupDiscountToExcel(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button59.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по скидкам по подгруппам товаров из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadSubgroupDiscountsFromLO()
        Else
            LoadSubgroupDiscountsFromExcel()
        End If
        MsgBox("Загрузка данных по скидкам по подгруппам продуктов из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadSubgroupDiscounts()
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub DataGridView14_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView14.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора скидки по группе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button64.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание скидки на товар
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountItem = New DiscountItem
        MyDiscountItem.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountItem.ShowDialog()
        '---загрузка данных
        LoadItemDiscounts()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView14.Rows.Count - 1
            If Trim(DataGridView14.Item(1, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView14.CurrentCell = DataGridView14.Item(1, i)
            End If
        Next
        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button65.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение скидки на товар
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountItem = New DiscountItem
        MyDiscountItem.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductID = Trim(Me.DataGridView14.SelectedRows.Item(0).Cells(1).Value)
        MyDiscountItem.ShowDialog()
        '---загрузка данных
        LoadItemDiscounts()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView14.Rows.Count - 1
            If Trim(DataGridView14.Item(1, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView14.CurrentCell = DataGridView14.Item(1, i)
            End If
        Next
        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button74.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Снятие фильтра окна скидок по подгруппе
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBS.Filter = ""
        Label21.BackColor = Color.White
        For i As Integer = 0 To DataGridView13.Columns.Count - 1
            DataGridView13.Columns(i).HeaderCell.Style.ForeColor = Color.Black
        Next
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор контекстного меню установка фильтра окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Declarations.MyFilterColumn = 1 Then
            MyBS.Filter = "GroupCode = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(1).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 2 Then
            MyBS.Filter = "Name = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(2).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 3 Then
            MyBS.Filter = "SubgroupCode = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(3).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 4 Then
            MyBS.Filter = "Expr1 = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(4).Value.ToString()) & "'"
        ElseIf Declarations.MyFilterColumn = 5 Then
            MyBS.Filter = "Discount = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(5).Value.ToString()) & "'"
        End If

        For i As Integer = 0 To DataGridView13.Columns.Count - 1
            If i = Declarations.MyFilterColumn Then
                DataGridView13.Columns(i).HeaderCell.Style.ForeColor = Color.Green
            Else
                DataGridView13.Columns(i).HeaderCell.Style.ForeColor = Color.Black
            End If
        Next
        Label21.BackColor = Color.Green
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор контекстного меню снятие фильтра окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBS.Filter = ""
        Label21.BackColor = Color.White
        For i As Integer = 0 To DataGridView13.Columns.Count - 1
            DataGridView13.Columns(i).HeaderCell.Style.ForeColor = Color.Black
        Next
    End Sub

    Private Sub Button67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button67.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel скидок по подгруппам для текущего клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadItemDiscountToLO(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        Else
            UploadItemDiscountToExcel(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button68.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по скидкам по товарам из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadItemDiscountsFromLO()
        Else
            LoadItemDiscountsFromExcel()
        End If
        MsgBox("Загрузка данных по скидкам по продуктам из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadItemDiscounts()
        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button66.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выбранной скидки на товар
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountItem "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView14.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---загрузка данных
        LoadItemDiscounts()
        CheckItemDiscountsButtons()
    End Sub

    Private Sub DataGridView15_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView15.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора согласованного ассортимента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button73.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание записи согласованного ассортимента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAgreedRange = New AgreedRange
        MyAgreedRange.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyAgreedRange.ShowDialog()
        '---загрузка данных
        LoadAgreedRange()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView15.Rows.Count - 1
            If Trim(DataGridView15.Item(1, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView15.CurrentCell = DataGridView15.Item(1, i)
            End If
        Next
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button72.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение скидки на товар
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAgreedRange = New AgreedRange
        MyAgreedRange.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductID = Trim(Me.DataGridView15.SelectedRows.Item(0).Cells(1).Value)
        MyAgreedRange.ShowDialog()
        '---загрузка данных
        LoadAgreedRange()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView15.Rows.Count - 1
            If Trim(DataGridView15.Item(1, i).Value.ToString) = Declarations.MyProductID Then
                DataGridView15.CurrentCell = DataGridView15.Item(1, i)
            End If
        Next
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button71.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выбранной записи согласованного ассортимента
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_AgreedRange "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView15.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---загрузка данных
        LoadAgreedRange()
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button70.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel записей о согласованном ассортименте для текущего клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadAgreedRangeToLO(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        Else
            UploadAgreedRangeToExcel(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text, _
            1)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button69.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по согласованному ассортименту из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadAgreedRangeFromLO()
        Else
            LoadAgreedRangeFromExcel()
        End If
        MsgBox("Загрузка данных по согласованному ассортименту из Excel завершена", MsgBoxStyle.Information, "Внимание!")
        '---загрузка данных
        LoadAgreedRange()
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button58.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel всей информации о скидках и согласованном ассортименте для текущего клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadFULLDiscountsAgreedRangeToLO(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text)
        Else
            UploadFULLDiscountsAgreedRangeToExcel(ComboBox1.SelectedValue, _
            ComboBox1.Text, _
            CDbl(IIf(Trim(TextBox3.Text) = "", 0, Trim(TextBox3.Text))), _
            TextBox4.Text)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel базового прайс листа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBasePrice = New BasePrice
        MyBasePrice.ShowDialog()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel индивидуального прайс листа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyIndPrice = New IndPrice
        MyIndPrice.ShowDialog()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации из Scala по всем таблицам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor

        MySQLStr = "exec spp_WEB_ALL_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Me.Cursor = Cursors.Default
        MsgBox("Загрузка информации из Scala произведена", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка из БД файлов в определенный каталог
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        FullUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("Полная выгрузка данных в каталог произведена.", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД файлов в определенный каталог - "ночная выгрузка"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        NightUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("Выгрузка данных в каталог произведена.", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации о доступности на складах из БД файлов в определенный каталог 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        AvailabilityUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("Выгрузка данных в каталог произведена.", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации о о продажах файлов в определенный каталог 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        SalesUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("Выгрузка данных в каталог произведена.", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка из БД файлов в определенный каталог с отправкой на WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If FullUploadToCatalog_WEB(1) = True Then
            MsgBox("Полная выгрузка данных на WEB произведена успешно.", MsgBoxStyle.Information, "Внимание!")
        Else
            MsgBox("Во время полной выгрузки данных на WEB были ошибки.", MsgBoxStyle.Information, "Внимание!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ночная (ежедневная) выгрузка из БД файлов в определенный каталог с отправкой на WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If DailyUploadToCatalog_WEB(1) = True Then
            MsgBox("Выгрузка данных на WEB произведена успешно.", MsgBoxStyle.Information, "Внимание!")
        Else
            MsgBox("Во время выгрузки данных на WEB были ошибки.", MsgBoxStyle.Information, "Внимание!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД информации о продажах - файлов в определенный каталог с отправкой на WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If SalesUploadToCatalog_WEB(1) = True Then
            MsgBox("Выгрузка данных на WEB произведена успешно.", MsgBoxStyle.Information, "Внимание!")
        Else
            MsgBox("Во время выгрузки данных на WEB были ошибки.", MsgBoxStyle.Information, "Внимание!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД информации о доступности на складах - файлов в определенный каталог с отправкой на WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If AvailabilityUploadToCatalog_WEB(1) = True Then
            MsgBox("Выгрузка данных на WEB произведена успешно.", MsgBoxStyle.Information, "Внимание!")
        Else
            MsgBox("Во время выгрузки данных на WEB были ошибки.", MsgBoxStyle.Information, "Внимание!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button75.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД информации для электронной торговли с Сен Гобеном в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadInfoToSaintGobain = New UploadInfoToSaintGobain
        MyUploadInfoToSaintGobain.ShowDialog()
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление картинок товаров,
        '// которые неверно корреспондируются с кодами товара поставщика в нашей БД.
        '// название файла соответствует коду товара поставщика 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDeletePictures = New DeletePictures
        MyDeletePictures.ShowDialog()
    End Sub

   
    Private Sub Button77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button77.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД картинок товаров
        '// в выбранный каталог
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadFilesToCatalog = New UploadFilesToCatalog
        MyUploadFilesToCatalog.ShowDialog()
    End Sub

    Private Sub Button76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button76.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка в БД картинок товаров
        '// картинки предварительно преобразовываются в .jpg
        '// название файла соответствует коду товара поставщика (как быть с теми, кто содержит недозволенные знаки?)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadFilesToDB = New UploadFilesToDB
        MyUploadFilesToDB.ShowDialog()
    End Sub

    Private Sub Button79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button79.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перенос в БД названий и описаний товаров
        '// полученных из WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyTransferNamesDescrToDB = New TransferNamesDescrToDB
        MyTransferNamesDescrToDB.ShowDialog()
    End Sub

    Private Sub Button80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button80.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна связывания картинок с кодами Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyMatchPictAndScalaCode = New MatchPictAndScalaCode
        MyMatchPictAndScalaCode.ShowDialog()
    End Sub

    Private Sub Button82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button82.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна удаления картинок из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDeletePictureFromDB = New DeletePictureFromDB
        MyDeletePictureFromDB.ShowDialog()
    End Sub

    Private Sub Button81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button81.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна загрузки одной картинки в БД
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyLoadOnePictToDB = New LoadOnePictToDB
        MyLoadOnePictToDB.ShowDialog()
    End Sub

    Private Sub Button83_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button83.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка в каталог
        '// данных с сервиса Шнейдер Электрик
        '// загружаются картинки (название - код товара поставщика) 
        '// названия, описания - в Excel в том же каталоге
        '////////////////////////////////////////////////////////////////////////////////

        MyDownloadInfoFromSE = New DownloadInfoFromSE
        MyDownloadInfoFromSE.ShowDialog()
    End Sub

    Private Sub Button85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button85.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации на сайт Magento только новой информации
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadDataToMagento = New UploadDataToMagento
        MyUploadDataToMagento.MyMode = 1
        MyUploadDataToMagento.ShowDialog()
    End Sub

    Private Sub Button84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button84.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Полная выгрузка информации на сайт Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadDataToMagento = New UploadDataToMagento
        MyUploadDataToMagento.MyMode = 0
        MyUploadDataToMagento.ShowDialog()
    End Sub

    Private Sub Button86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button86.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации о доступности на складах на сайт Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadAvailabilityToMagento = New UploadAvailabilityToMagento
        MyUploadAvailabilityToMagento.ShowDialog()
    End Sub

    Private Sub Button87_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button87.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление картинок на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadPicturesToMagento = New UploadPicturesToMagento
        MyUploadPicturesToMagento.ShowDialog()
    End Sub

    Private Sub Button88_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button88.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации по всем прайсам на сайт Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub Button89_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button89.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка незаблокированной номенклатуры из Scala 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCASH_FullUpload = New CASH_FullUpload
        MyCASH_FullUpload.ShowDialog()

    End Sub

    Private Sub Button90_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button90.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка определенных обобщенных названий 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCASH_CustomUpload = New CASH_CustomUpload
        MyCASH_CustomUpload.ShowDialog()
    End Sub

    Private Sub Button91_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button91.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нахождение запаса по Скальскому коду  
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "SELECT tbl_WEB_Items.GroupCode, ISNULL(tbl_WEB_ItemSubGroup.SubgroupID, N'') AS SubgroupID "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Items LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup ON tbl_WEB_Items.GroupCode = tbl_WEB_ItemSubGroup.GroupCode AND "
        MySQLStr = MySQLStr & "tbl_WEB_Items.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Items.Code = N'" & Trim(TextBox5.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Запас " & Trim(TextBox5.Text) & " не найден. ", MsgBoxStyle.Critical, "Внимание!")
        Else
            '----- Выставление группы запаса
            For i As Integer = 0 To DataGridView7.Rows.Count - 1
                If DataGridView7.Item(0, i).Value = Trim(Declarations.MyRec.Fields("GroupCode").Value) Then
                    DataGridView7.CurrentCell = DataGridView7.Item(1, i)
                    Exit For
                End If
            Next

            '----- Выставление подгруппы запаса или нахожддение запаса
            If Trim(Declarations.MyRec.Fields("GroupCode").Value).Equals("") Then
                '-----товар не в подгруппе - находим его в товарах без подгруппы
                For i As Integer = 0 To DataGridView9.Rows.Count - 1
                    If DataGridView9.Item(1, i).Value = Trim(TextBox5.Text) Then
                        DataGridView9.CurrentCell = DataGridView9.Item(1, i)
                        Exit For
                    End If
                Next
            Else
                '-----товар в подгруппе
                '-----сначала находим подгруппу и выставляем
                For i As Integer = 0 To DataGridView8.Rows.Count - 1
                    If DataGridView8.Item(0, i).Value = Trim(Declarations.MyRec.Fields("SubgroupID").Value) Then
                        DataGridView8.CurrentCell = DataGridView8.Item(1, i)
                        Exit For
                    End If
                Next

                '-----находим товар в товарах с подгруппой
                For i As Integer = 0 To DataGridView10.Rows.Count - 1
                    If DataGridView10.Item(1, i).Value = Trim(TextBox5.Text) Then
                        DataGridView10.CurrentCell = DataGridView10.Item(1, i)
                        Exit For
                    End If
                Next
            End If
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button93_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button93.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка в Excel информации о длине, ширине, высоте и весе товара  
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadItemDimToLO()
        Else
            UploadItemDimToExcel()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button92_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button92.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Excel информации о длине, ширине, высоте и весе товара  
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadItemDimFromLO()
        Else
            LoadItemDimFromExcel()
        End If
        MsgBox("Загрузка данных по продуктам завершена", MsgBoxStyle.Information, "Внимание!")
        Cursor = Cursors.Default
    End Sub

    Private Sub Button94_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button94.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка в каталог
        '// данных с сервиса ABB
        '// загружаются картинки (название - код товара поставщика) 
        '// названия, описания - в Excel в том же каталоге
        '////////////////////////////////////////////////////////////////////////////////

        MyDownloadInfoFromABB = New DownloadInfoFromABB
        MyDownloadInfoFromABB.ShowDialog()
    End Sub

End Class


