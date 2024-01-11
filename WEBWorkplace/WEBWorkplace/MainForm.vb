
Public Class MainForm
    Public MyBS As New BindingSource()

    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Tab �������������
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
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��������
        '// ��� ��������� �������� ������� ������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        Select Case sender.selectedtab.text
            Case "������"
                LoadCities()
                CheckCitiesButtons()
            Case "�������������"
                LoadManufacturers()
                CheckManufacturersButtons()
            Case "��������"
                LoadSalesmans()
                CheckSalesmansButtons()
            Case "������ ������"
                LoadProductGroup()
                CheckProductGroupButtons()
            Case "��������� ������"
                LoadProductSubgroup()
                CheckProductSubGroupButtons()
            Case "������"
                LoadProducts()
                CheckProductButtons()
            Case "�������"
                LoadCustomers()
                CheckCustomerButtons()
            Case "������ � ������. �����������"
                LoadDiscountsHeader()
                'Case "��������"
            Case Else
        End Select
    End Sub

    Private Sub LoadCities()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ �������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

        DataGridView1.Columns(0).HeaderText = "��� ������"
        DataGridView1.Columns(0).Width = 200
        DataGridView1.Columns(1).HeaderText = "�������� ������"
        DataGridView1.Columns(1).Width = 500

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadManufacturers()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ��������������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

        DataGridView2.Columns(0).HeaderText = "��� �������������"
        DataGridView2.Columns(0).Width = 100
        DataGridView2.Columns(1).HeaderText = "�������� �������������"
        DataGridView2.Columns(1).Width = 290
        DataGridView2.Columns(2).HeaderText = "���������� �������� �������������"
        DataGridView2.Columns(2).Width = 290
        DataGridView2.Columns(3).HeaderText = "��������� ����"
        DataGridView2.Columns(3).Width = 250
        DataGridView2.Columns(4).HeaderText = "������ Scala"
        DataGridView2.Columns(4).Width = 50
        DataGridView2.Columns(5).HeaderText = "������ WEB"
        DataGridView2.Columns(5).Width = 50

        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadSalesmans()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_WEB_Salesmans.Code, tbl_WEB_Salesmans.Name, tbl_WEB_Salesmans.Email, ISNULL(tbl_WEB_Cities.Name, '') AS City, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_WEB_Salesmans.OfficeLeader = 0 THEN '' ELSE '��' END AS OfficeLeader, CASE WHEN tbl_WEB_Salesmans.OnDuty = 0 THEN '' ELSE '��' END AS OnDuty, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_WEB_Salesmans.IsActive = 0 THEN '' ELSE '��������' END AS IsActive, tbl_WEB_Salesmans.Rezerv1, "
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

        DataGridView3.Columns(0).HeaderText = "��� ��������"
        DataGridView3.Columns(0).Width = 50
        DataGridView3.Columns(1).HeaderText = "��� ��������"
        DataGridView3.Columns(1).Width = 150
        DataGridView3.Columns(2).HeaderText = "E-mail ��������"
        DataGridView3.Columns(2).Width = 200
        DataGridView3.Columns(3).HeaderText = "����� ��������"
        DataGridView3.Columns(3).Width = 150
        DataGridView3.Columns(4).HeaderText = "����� ����� ��� �� WEB � ������"
        DataGridView3.Columns(4).Width = 50
        DataGridView3.Columns(5).HeaderText = "����� ���"
        DataGridView3.Columns(5).Width = 50
        DataGridView3.Columns(6).HeaderText = "����� ���"
        DataGridView3.Columns(6).Width = 50
        DataGridView3.Columns(7).HeaderText = "��������� ���� 1"
        DataGridView3.Columns(7).Width = 100
        DataGridView3.Columns(8).HeaderText = "��������� ���� 2"
        DataGridView3.Columns(8).Width = 100
        DataGridView3.Columns(9).HeaderText = "������ ��"
        DataGridView3.Columns(9).Width = 50
        DataGridView3.Columns(10).HeaderText = "������ WEB"
        DataGridView3.Columns(10).Width = 50
        DataGridView3.Columns(11).HeaderText = "������ Scala"
        DataGridView3.Columns(11).Width = 50

        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductGroup()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ����� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

        DataGridView4.Columns(0).HeaderText = "��� ������"
        DataGridView4.Columns(0).Width = 100
        DataGridView4.Columns(1).HeaderText = "�������� �� Scala"
        DataGridView4.Columns(1).Width = 420
        DataGridView4.Columns(2).HeaderText = "�������� ��� WEB"
        DataGridView4.Columns(2).Width = 420
        DataGridView4.Columns(3).HeaderText = "������ Scala"
        DataGridView4.Columns(3).Width = 50
        DataGridView4.Columns(4).HeaderText = "������ WEB"
        DataGridView4.Columns(4).Width = 50

        DataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroup()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� �������� ��������� ������ ����� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---��� ������� ������, � ���� ��������� - �� ��������� ������ ������
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

        DataGridView5.Columns(0).HeaderText = "��� ������"
        DataGridView5.Columns(0).Width = 100
        DataGridView5.Columns(1).HeaderText = "�������� �� Scala"
        DataGridView5.Columns(1).Width = 420
        DataGridView5.Columns(2).HeaderText = "�������� ��� WEB"
        DataGridView5.Columns(2).Width = 420
        DataGridView5.Columns(3).HeaderText = "������ Scala"
        DataGridView5.Columns(3).Width = 50
        DataGridView5.Columns(4).HeaderText = "������ WEB"
        DataGridView5.Columns(4).Width = 50

        DataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroupDetail()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ������� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView6.Columns(0).HeaderText = "ID ��� ������"
            DataGridView6.Columns(0).Width = 70
            DataGridView6.Columns(1).HeaderText = "��� ��� ������"
            DataGridView6.Columns(1).Width = 70
            DataGridView6.Columns(2).HeaderText = "��� ������"
            DataGridView6.Columns(2).Width = 70
            DataGridView6.Columns(3).HeaderText = "�������� ���������"
            DataGridView6.Columns(3).Width = 320
            DataGridView6.Columns(4).HeaderText = "�������� ���������"
            DataGridView6.Columns(4).Width = 320
            DataGridView6.Columns(5).HeaderText = "��������� ����"
            DataGridView6.Columns(5).Width = 100
            DataGridView6.Columns(6).HeaderText = "������ Scala"
            DataGridView6.Columns(6).Width = 50
            DataGridView6.Columns(7).HeaderText = "������ WEB"
            DataGridView6.Columns(7).Width = 50
        End If
        DataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProducts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������� ������ ����� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---��� ������� ������, � ���� ��������� - �� ��������� ������ ������
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

        DataGridView7.Columns(0).HeaderText = "��� ������"
        DataGridView7.Columns(0).Width = 100
        DataGridView7.Columns(1).HeaderText = "�������� �� Scala"
        DataGridView7.Columns(1).Width = 390

        DataGridView7.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProductSubgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������� ������ �������� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView8.Columns(0).HeaderText = "ID ���������"
            DataGridView8.Columns(0).Width = 0
            DataGridView8.Columns(0).Visible = False
            DataGridView8.Columns(1).HeaderText = "��� ���������"
            DataGridView8.Columns(1).Width = 100
            DataGridView8.Columns(2).HeaderText = "�������� ���������"
            DataGridView8.Columns(2).Width = 390

        End If
        DataGridView8.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub

    Private Sub LoadProduct_IN_Subgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������� ������ �������, ������������� � ��������� ��������� �������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView10.Columns(0).HeaderText = "���� ����"
            DataGridView10.Columns(0).Width = 35
            DataGridView10.Columns(1).HeaderText = "��� ������ � Scala"
            DataGridView10.Columns(1).Width = 100
            DataGridView10.Columns(2).HeaderText = "��� ������ � Scala"
            DataGridView10.Columns(2).Width = 300
            DataGridView10.Columns(3).HeaderText = "�������������"
            DataGridView10.Columns(3).Width = 150
            DataGridView10.Columns(4).HeaderText = "��� ������ �������������"
            DataGridView10.Columns(4).Width = 150
            DataGridView10.Columns(5).HeaderText = "������"
            DataGridView10.Columns(5).Width = 100
            DataGridView10.Columns(6).HeaderText = "��� ������ ��� WEB"
            DataGridView10.Columns(6).Width = 300
            DataGridView10.Columns(7).HeaderText = "�������� ������"
            DataGridView10.Columns(7).Width = 300
            DataGridView10.Columns(8).HeaderText = "��������� �����������"
            DataGridView10.Columns(8).Width = 70
            DataGridView10.Columns(9).HeaderText = "������� ���������"
            DataGridView10.Columns(9).Width = 70
            DataGridView10.Columns(10).HeaderText = "��������� ����2"
            DataGridView10.Columns(10).Width = 100
            DataGridView10.Columns(11).HeaderText = "������ Scala"
            DataGridView10.Columns(11).Width = 40
            DataGridView10.Columns(12).HeaderText = "������ WEB"
            DataGridView10.Columns(12).Width = 40

        End If
        DataGridView10.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView10.MultiSelect = True
    End Sub

    Private Sub LoadProduct_NO_Subgroup_IN_Items()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������� ������ �������, �� ������������� �� � ����� ��������� �������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView9.Columns(0).HeaderText = "���� ����"
            DataGridView9.Columns(0).Width = 40
            DataGridView9.Columns(1).HeaderText = "��� ������ � Scala"
            DataGridView9.Columns(1).Width = 100
            DataGridView9.Columns(2).HeaderText = "��� ������ � Scala"
            DataGridView9.Columns(2).Width = 300
            DataGridView9.Columns(3).HeaderText = "�������������"
            DataGridView9.Columns(3).Width = 150
            DataGridView9.Columns(4).HeaderText = "��� ������ �������������"
            DataGridView9.Columns(4).Width = 150
            DataGridView9.Columns(5).HeaderText = "������"
            DataGridView9.Columns(5).Width = 100
            DataGridView9.Columns(6).HeaderText = "��� ������ ��� WEB"
            DataGridView9.Columns(6).Width = 300
            DataGridView9.Columns(7).HeaderText = "�������� ������"
            DataGridView9.Columns(7).Width = 300
            DataGridView9.Columns(8).HeaderText = "��������� �����������"
            DataGridView9.Columns(8).Width = 70
            DataGridView9.Columns(9).HeaderText = "������� ���������"
            DataGridView9.Columns(9).Width = 70
            DataGridView9.Columns(10).HeaderText = "��������� ����"
            DataGridView9.Columns(10).Width = 100
            DataGridView9.Columns(11).HeaderText = "������ Scala"
            DataGridView9.Columns(11).Width = 40
            DataGridView9.Columns(12).HeaderText = "������ WEB"
            DataGridView9.Columns(12).Width = 40

        End If
        DataGridView9.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView9.MultiSelect = True
    End Sub

    Private Sub LoadCustomers()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ��������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        MySQLStr = "SELECT Code, Name, Address, Discount, Case WHEN WorkOverWEB = 1 THEN '��' ELSE '' END as WorkOverWEB, BasePrice "
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

        DataGridView11.Columns(0).HeaderText = "��� �������"
        DataGridView11.Columns(0).Width = 100
        DataGridView11.Columns(1).HeaderText = "�������� �������"
        DataGridView11.Columns(1).Width = 250
        DataGridView11.Columns(2).HeaderText = "�����"
        DataGridView11.Columns(2).Width = 520
        DataGridView11.Columns(3).HeaderText = "����� ������"
        DataGridView11.Columns(3).Width = 60
        DataGridView11.Columns(4).HeaderText = "�������� ����� WEB"
        DataGridView11.Columns(4).Width = 60
        DataGridView11.Columns(5).HeaderText = "������� �����"
        DataGridView11.Columns(5).Width = 60

        DataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadDiscountsHeader()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ���� ������ ��������, ���������� ����� WEB
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ ��������
        Dim MyDs As New DataSet

        '---------------������ ��������
        MySQLStr = "SELECT Code, LTRIM(RTRIM(LTRIM(RTRIM(Code)) + ' ' + LTRIM(RTRIM(Name)))) AS Name "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Clients "
        MySQLStr = MySQLStr & "WHERE (WorkOverWEB = 1) "
        MySQLStr = MySQLStr & "ORDER BY Name "

        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "Code"   '��� �� ��� ����� ���������
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
        '// �������� � ��������� ������� ������ ��� �������
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
        '// �������� � ��������� ������� ������ ��� ��������������
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
        '// �������� � ��������� ������� ������ ��� ���������
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
        '// �������� � ��������� ������� ������ ��� ����� ���������
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
        '// �������� � ��������� ������� ������ ��� �������� ���������
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
        '// �������� � ��������� ������� ������ ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '----------������ ����������� �/�� ���������
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

        '---------������ �������������� �������
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

        '----------������ �������� � Excel
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
        '// �������� � ��������� ������� ������ ��� �������
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
        '// ��������� �������� � Tab �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim RemFlag As Integer          '���� - ����� �� ������� ������

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
            MsgBox("������� ������ ������ �������, ��� ��� ���� ������ �� ���������, ����������� �� ������ �����.", MsgBoxStyle.Critical, "��������!")
        Else
            MySQLStr = "DELETE FROM tbl_WEB_Cities "
            MySQLStr = MySQLStr & "WHERE (ID = " & Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---�������� ������
            LoadCities()
            CheckCitiesButtons()
        End If

    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCity = New City
        MyCity.StartParam = "Create"
        Declarations.MyCityID = 0
        MyCity.ShowDialog()
        '---�������� ������
        LoadCities()
        '---������� ������� ������� ���������
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
        '// �������������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCity = New City
        MyCity.StartParam = "Edit"
        Declarations.MyCityID = Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyCity.ShowDialog()
        '---�������� ������
        LoadCities()
        '---������� ������� ������� ����������
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
        '// ��������� ����� �������������� � ����������� �� �������
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
        '// �������� ���������� �� ����������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Manufacturers_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("�������� ���������� �� �������������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
        LoadManufacturers()
        CheckManufacturersButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyManufacturer = New Manufacturer
        Declarations.MyManufacturerID = Me.DataGridView2.SelectedRows.Item(0).Cells(0).Value
        MyManufacturer.ShowDialog()
        '---�������� ������
        LoadManufacturers()
        '---������� ������� ������� ���������
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
        '// ����� ������ �������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckManufacturersButtons()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckCitiesButtons()
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������������� � Excel
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
        '// �������� �������������� � Excel
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
        '// �������� ���������� �� �������������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadManufacturersFromLO()
        Else
            LoadManufacturersFromExcel()
        End If
        MsgBox("�������� ������ �� �������������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadManufacturers()
        CheckManufacturersButtons()
    End Sub

    Private Sub DataGridView3_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ��������� � ����������� �� �������
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
        '// ����� ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSalesmansButtons()
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ��������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Salesmans_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("�������� ���������� �� ��������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
        LoadSalesmans()
        CheckSalesmansButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySalesman = New Salesman
        Declarations.MySalesmanID = Me.DataGridView3.SelectedRows.Item(0).Cells(0).Value
        MySalesman.ShowDialog()
        '---�������� ������
        LoadSalesmans()
        '---������� ������� ������� ����������
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
        '// �������� ��������� � Excel
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
        '// �������� ��������� � Excel
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
        '// �������� ���������� �� ��������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadSalesmansFromLO()
        Else
            LoadSalesmansFromExcel()
        End If
        MsgBox("�������� ������ �� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadSalesmans()
        CheckSalesmansButtons()
    End Sub

    Private Sub DataGridView4_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ����� ��������� � ����������� �� �������
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
        '// ����� ������ ������ ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProductGroupButtons()
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ������� ��������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_ItemGroups_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("�������� ���������� �� ������� ��������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
        LoadProductGroup()
        CheckProductGroupButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ������ ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductGroup = New ProductGroup
        Declarations.MyProductGroupID = Me.DataGridView4.SelectedRows.Item(0).Cells(0).Value
        MyProductGroup.ShowDialog()
        '---�������� ������
        LoadProductGroup()
        '---������� ������� ������� ����������
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
        '// �������� ����� ��������� � Excel
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
        '// �������� ����� ��������� � Excel
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
        '// �������� ���������� �� ������� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductGroupsFromLO()
        Else
            LoadProductGroupsFromExcel()
        End If
        MsgBox("�������� ������ �� ������� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadProductGroup()
        CheckProductGroupButtons()
    End Sub

    Private Sub DataGridView5_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ����� ��������� � ����������� �� �������
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
        '// ����� ������ ������ �������� � �������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProductSubgroupDetail()
    End Sub

    Private Sub DataGridView6_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView6.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� �������� ��������� � ����������� �� �������
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
        '// ����� ������ ��������� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProductSubGroupButtons()
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� ��������� � Excel (������ ��� ������� ������)
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
        '// �������� �������� ��������� � Excel (��� ���� �����)
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
        '// �������� �������� ��������� � Excel (��� ���� �����)
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
        '// �������� ���������� �� ���������� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductSubGroupsFromLO()
        Else
            LoadProductSubGroupsFromExcel()
        End If
        MsgBox("�������� ������ �� ���������� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadProductSubgroup()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductSubGroup = New ProductSubGroup
        MyProductSubGroup.StartParam = "Create"
        Declarations.MyProductGroupID = Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductSubGroupID = "N"
        MyProductSubGroup.ShowDialog()
        '---�������� ������
        LoadProductSubgroupDetail()
        '---������� ������� ������� ���������
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
        '// �������������� ��������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProductSubGroup = New ProductSubGroup
        MyProductSubGroup.StartParam = "Edit"
        Declarations.MyProductGroupID = Trim(Me.DataGridView5.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductSubGroupID = Trim(Me.DataGridView6.SelectedRows.Item(0).Cells(1).Value)
        MyProductSubGroup.ShowDialog()
        '---�������� ������
        LoadProductSubgroupDetail()
        '---������� ������� ������� ���������
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
        '// �������� ���������� ��������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim RemFlag As Integer          '���� - ����� �� ������� ������

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
            MsgBox("������� ��������� ������ �������, ��� ��� ���� ������, �������� � ������ ���������.", MsgBoxStyle.Critical, "��������!")
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
            '---�������� ������
            LoadProductSubgroupDetail()
            CheckProductSubGroupButtons()
        End If
    End Sub

    Private Sub DataGridView7_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView7.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������ �������� � �������� ������
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
        '// ����� ������ ��������� �������� � �������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProduct_IN_Subgroup_IN_Items()
        CheckProductButtons()
    End Sub

    Private Sub DataGridView9_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView9.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ��������� � ����������� �� �������
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
        '// ��������� ����� ��������� � ����������� �� �������
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
        '// �������� ���������� �� ������� ��������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Items_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("�������� ���������� �� ��������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
        LoadProducts()
        CheckProductButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �� ���������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim i As Integer
        Dim MySQLStr As String

        If DataGridView10.SelectedRows.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            For i = 0 To DataGridView10.SelectedRows.Count - 1
                '--------���������� ���������� � ������
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET SubGroupCode = N'', "
                MySQLStr = MySQLStr & "RMStatus = 2, "
                MySQLStr = MySQLStr & "WEBStatus = 2 "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView10.SelectedRows.Item(i).Cells(1).Value & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '------���������� ���������� � �������
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
        '// ���������� ������� � ���������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer
        Dim MySQLStr As String

        If DataGridView9.SelectedRows.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            For i = 0 To DataGridView9.SelectedRows.Count - 1
                '------���������� ���������� � ������
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET SubGroupCode = N'" & Me.DataGridView8.SelectedRows.Item(0).Cells(1).Value & "', "
                MySQLStr = MySQLStr & "RMStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE 1 END, "
                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN ScalaStatus = 2 THEN 2 ELSE 1 END "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & DataGridView9.SelectedRows.Item(i).Cells(1).Value & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '------���������� ���������� � �������
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
        '// �������� ��������� � Excel (��� ������ ������ ��� ������� ������)
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
        '// �������� ��������� � Excel (��� ������ ������ ��� ������� ������, ��� ������� ���������)
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
        '// �������� ��������� � Excel (��� ������� ������, �� �������� �� � ���� ���������)
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
        '// �������� ��������� � Excel (��� ������)
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
        '// �������� ��������� � Excel (��� ������ ��� ��������)
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
        '// �������� ��������� � Excel (��� ������) � �������� ��������
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
        '// �������� ���������� �� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadProductsFromLO()
        Else
            LoadProductsFromExcel()
        End If

        MsgBox("�������� ������ �� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        CheckProductButtons()
        Cursor = Cursors.Default
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProduct = New Product
        MyProduct.StartParam = "NotInGroup"
        MyProduct.MyGroup = Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductID = Trim(Me.DataGridView9.SelectedRows.Item(0).Cells(1).Value)
        MyProduct.ShowDialog()
        '---�������� ������
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        '---������� ������� ������� ���������
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
        '// �������������� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProduct = New Product
        MyProduct.StartParam = "InGroup"
        MyProduct.MyGroup = Trim(Me.DataGridView7.SelectedRows.Item(0).Cells(0).Value)
        Declarations.MyProductID = Trim(Me.DataGridView10.SelectedRows.Item(0).Cells(1).Value)
        MyProduct.ShowDialog()
        '---�������� ������
        LoadProduct_IN_Subgroup_IN_Items()
        LoadProduct_NO_Subgroup_IN_Items()
        '---������� ������� ������� ���������
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
        '// ��������� ����� �������� � ����������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView11.Rows(e.RowIndex)
        If row.Cells(4).Value = "��" Then
            row.DefaultCellStyle.BackColor = Color.LightGreen
        Else
            row.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView11_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView11.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������� � �������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckCustomerButtons()
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���� ���������� �� �������� ����������� � ��������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("���������� ������ �������� ������", MsgBoxStyle.OkOnly, "��������!")
            TextBox1.Select()
        Else
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button51.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor
        MySQLStr = "exec spp_WEB_Clients_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MsgBox("�������� ���������� �� �������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
        LoadCustomers()
        CheckCustomerButtons()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button52.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomer = New Customer
        Declarations.MyCustomerID = Trim(Me.DataGridView11.SelectedRows.Item(0).Cells(0).Value)
        MyCustomer.ShowDialog()
        '---�������� ������
        LoadCustomers()
        '---������� ������� ������� ���������
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
        '// ����� ��������
        '// ��� ��������� �������� ������� ������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        Select Case sender.selectedtab.text
            Case "������ �� ������ ���������"
                LoadGroupDiscount()
                CheckGroupDiscountsButtons()
            Case "������ �� ��������� ���������"
                LoadSubgroupDiscounts()
                CheckSubgroupDiscountsButtons()
            Case "������ �� ��������"
                LoadItemDiscounts()
                CheckItemDiscountsButtons()
            Case "������������� �����������"
                LoadAgreedRange()
                CheckAgreedRangeButtons()
            Case Else
        End Select
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ �������, ��� �������� �������� ������
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

        If TabControl2.SelectedTab.Text = "������ �� ������ ���������" Then
            LoadGroupDiscount()
            CheckGroupDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "������ �� ��������� ���������" Then
            LoadSubgroupDiscounts()
            CheckSubgroupDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "������ �� ��������" Then
            LoadItemDiscounts()
            CheckItemDiscountsButtons()
        ElseIf TabControl2.SelectedTab.Text = "������������� �����������" Then
            LoadAgreedRange()
            CheckAgreedRangeButtons()
        Else

        End If
    End Sub

    Private Sub LoadGroupDiscount()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ������ �� ������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView12.Columns(0).HeaderText = "ID ������ �������"
            DataGridView12.Columns(0).Width = 0
            DataGridView12.Columns(0).Visible = False
            DataGridView12.Columns(1).HeaderText = "��� ������ �������"
            DataGridView12.Columns(1).Width = 150
            DataGridView12.Columns(2).HeaderText = "�������� ������ �������"
            DataGridView12.Columns(2).Width = 800
            DataGridView12.Columns(3).HeaderText = "������ (%)"
            DataGridView12.Columns(3).Width = 100
        End If

        DataGridView12.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadSubgroupDiscounts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ������ �� ���������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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
                '---������
                MyBS.DataSource = MyDs
                MyBS.DataMember = MyDs.Tables(0).TableName
                DataGridView13.DataSource = MyBS
                '---����� �������

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            DataGridView13.Columns(0).HeaderText = "ID ������ �������"
            DataGridView13.Columns(0).Width = 0
            DataGridView13.Columns(0).Visible = False
            DataGridView13.Columns(1).HeaderText = "��� ������ �������"
            DataGridView13.Columns(1).Width = 80
            DataGridView13.Columns(2).HeaderText = "�������� ������ �������"
            DataGridView13.Columns(2).Width = 400
            DataGridView13.Columns(3).HeaderText = "��� ��������� �������"
            DataGridView13.Columns(3).Width = 80
            DataGridView13.Columns(4).HeaderText = "�������� ��������� �������"
            DataGridView13.Columns(4).Width = 400
            DataGridView13.Columns(5).HeaderText = "������ (%)"
            DataGridView13.Columns(5).Width = 90
        End If

        DataGridView13.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadItemDiscounts()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ ������ �� �������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView14.Columns(0).HeaderText = "ID ������"
            DataGridView14.Columns(0).Width = 0
            DataGridView14.Columns(0).Visible = False
            DataGridView14.Columns(1).HeaderText = "��� ������"
            DataGridView14.Columns(1).Width = 150
            DataGridView14.Columns(2).HeaderText = "�������� ������"
            DataGridView14.Columns(2).Width = 800
            DataGridView14.Columns(3).HeaderText = "������ (%)"
            DataGridView14.Columns(3).Width = 100
        End If

        DataGridView14.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub LoadAgreedRange()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � ���� ������ �������������� ������������ (��� ���)
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
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

            DataGridView15.Columns(0).HeaderText = "ID ������"
            DataGridView15.Columns(0).Width = 0
            DataGridView15.Columns(0).Visible = False
            DataGridView15.Columns(1).HeaderText = "��� ������"
            DataGridView15.Columns(1).Width = 150
            DataGridView15.Columns(2).HeaderText = "�������� ������"
            DataGridView15.Columns(2).Width = 740
            DataGridView15.Columns(3).HeaderText = "���� (��� ���)"
            DataGridView15.Columns(3).Width = 80
            DataGridView15.Columns(4).HeaderText = "������"
            DataGridView15.Columns(4).Width = 80
        End If

        DataGridView15.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub DataGridView12_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView12.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckGroupDiscountsButtons()
    End Sub

    Private Sub CheckGroupDiscountsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ��������� ������� ������ ��� ������ �� ������
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
        '// �������� � ��������� ������� ������ ��� ������ �� ���������
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
        '// �������� � ��������� ������� ������ ��� ������ �� ������
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
        '// �������� � ��������� ������� ������ ��� �������������� ������������
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
        '// �������� ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountGroup = New DiscountGroup
        MyDiscountGroup.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountGroup.ShowDialog()
        '---�������� ������
        LoadGroupDiscount()
        '---������� ������� ������� ���������
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
        '// �������������� ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountGroup = New DiscountGroup
        MyDiscountGroup.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductGroupID = Trim(Me.DataGridView12.SelectedRows.Item(0).Cells(1).Value)
        MyDiscountGroup.ShowDialog()
        '---�������� ������
        LoadGroupDiscount()
        '---������� ������� ������� ���������
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
        '// �������� ��������� ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountGroup "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView12.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---�������� ������
        LoadGroupDiscount()
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub Button56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button56.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ������ �� ������� ��� �������� �������
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
        '// �������� ���������� �� ������� �� ������� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadGroupDiscountsFromLO()
        Else
            LoadGroupDiscountsFromExcel()
        End If
        MsgBox("�������� ������ �� ������� �� ������� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadGroupDiscount()
        CheckGroupDiscountsButtons()
    End Sub

    Private Sub DataGridView13_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView13.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���� ����������� �������
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
        '// ����� ������ ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button63.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountSubgroup = New DiscountSubgroup
        MyDiscountSubgroup.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountSubgroup.ShowDialog()
        '---�������� ������
        LoadSubgroupDiscounts()
        '---������� ������� ������� ���������
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
        '// ��������� ������ �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountSubgroup = New DiscountSubgroup
        MyDiscountSubgroup.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductGroupID = Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(1).Value)
        Declarations.MyProductSubGroupID = Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(3).Value)
        MyDiscountSubgroup.ShowDialog()
        '---�������� ������
        LoadSubgroupDiscounts()
        '---������� ������� ������� ���������
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
        '// �������� ��������� ������ �� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountSubgroup "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView13.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---�������� ������
        LoadSubgroupDiscounts()
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub Button60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button60.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ������ �� ���������� ��� �������� �������
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
        '// �������� ���������� �� ������� �� ���������� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadSubgroupDiscountsFromLO()
        Else
            LoadSubgroupDiscountsFromExcel()
        End If
        MsgBox("�������� ������ �� ������� �� ���������� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadSubgroupDiscounts()
        CheckSubgroupDiscountsButtons()
    End Sub

    Private Sub DataGridView14_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView14.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������ �� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button64.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountItem = New DiscountItem
        MyDiscountItem.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyDiscountItem.ShowDialog()
        '---�������� ������
        LoadItemDiscounts()
        '---������� ������� ������� ���������
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
        '// ��������� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDiscountItem = New DiscountItem
        MyDiscountItem.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductID = Trim(Me.DataGridView14.SelectedRows.Item(0).Cells(1).Value)
        MyDiscountItem.ShowDialog()
        '---�������� ������
        LoadItemDiscounts()
        '---������� ������� ������� ���������
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
        '// ������ ������� ���� ������ �� ���������
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
        '// ����� ������������ ���� ��������� ������� ����
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
        '// ����� ������������ ���� ������ ������� ����
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
        '// �������� � Excel ������ �� ���������� ��� �������� �������
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
        '// �������� ���������� �� ������� �� ������� �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadItemDiscountsFromLO()
        Else
            LoadItemDiscountsFromExcel()
        End If
        MsgBox("�������� ������ �� ������� �� ��������� �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadItemDiscounts()
        CheckItemDiscountsButtons()
    End Sub

    Private Sub Button66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button66.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_DiscountItem "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView14.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---�������� ������
        LoadItemDiscounts()
        CheckItemDiscountsButtons()
    End Sub

    Private Sub DataGridView15_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView15.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ �������������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button73.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �������������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAgreedRange = New AgreedRange
        MyAgreedRange.StartParam = "Create"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        MyAgreedRange.ShowDialog()
        '---�������� ������
        LoadAgreedRange()
        '---������� ������� ������� ���������
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
        '// ��������� ������ �� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAgreedRange = New AgreedRange
        MyAgreedRange.StartParam = "Edit"
        Declarations.MyCustomerID = Trim(ComboBox1.SelectedValue)
        Declarations.MyProductID = Trim(Me.DataGridView15.SelectedRows.Item(0).Cells(1).Value)
        MyAgreedRange.ShowDialog()
        '---�������� ������
        LoadAgreedRange()
        '---������� ������� ������� ���������
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
        '// �������� ��������� ������ �������������� ������������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_AgreedRange "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView15.SelectedRows.Item(0).Cells(0).Value.ToString) & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '---�������� ������
        LoadAgreedRange()
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button70.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ������� � ������������� ������������ ��� �������� �������
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
        '// �������� ���������� �� �������������� ������������ �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadAgreedRangeFromLO()
        Else
            LoadAgreedRangeFromExcel()
        End If
        MsgBox("�������� ������ �� �������������� ������������ �� Excel ���������", MsgBoxStyle.Information, "��������!")
        '---�������� ������
        LoadAgreedRange()
        CheckAgreedRangeButtons()
    End Sub

    Private Sub Button58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button58.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ���� ���������� � ������� � ������������� ������������ ��� �������� �������
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
        '// �������� � Excel �������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyBasePrice = New BasePrice
        MyBasePrice.ShowDialog()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � Excel ��������������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyIndPrice = New IndPrice
        MyIndPrice.ShowDialog()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �� Scala �� ���� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        Me.Cursor = Cursors.WaitCursor

        MySQLStr = "exec spp_WEB_ALL_FromScala "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Me.Cursor = Cursors.Default
        MsgBox("�������� ���������� �� Scala �����������", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� �� �� ������ � ������������ �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        FullUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("������ �������� ������ � ������� �����������.", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� �� ������ � ������������ ������� - "������ ��������"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        NightUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("�������� ������ � ������� �����������.", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � ����������� �� ������� �� �� ������ � ������������ ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        AvailabilityUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("�������� ������ � ������� �����������.", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � � �������� ������ � ������������ ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        SalesUploadToCatalog(0)
        Me.Cursor = Cursors.Default
        MsgBox("�������� ������ � ������� �����������.", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� �� �� ������ � ������������ ������� � ��������� �� WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If FullUploadToCatalog_WEB(1) = True Then
            MsgBox("������ �������� ������ �� WEB ����������� �������.", MsgBoxStyle.Information, "��������!")
        Else
            MsgBox("�� ����� ������ �������� ������ �� WEB ���� ������.", MsgBoxStyle.Information, "��������!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ (����������) �������� �� �� ������ � ������������ ������� � ��������� �� WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If DailyUploadToCatalog_WEB(1) = True Then
            MsgBox("�������� ������ �� WEB ����������� �������.", MsgBoxStyle.Information, "��������!")
        Else
            MsgBox("�� ����� �������� ������ �� WEB ���� ������.", MsgBoxStyle.Information, "��������!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� �� ���������� � �������� - ������ � ������������ ������� � ��������� �� WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If SalesUploadToCatalog_WEB(1) = True Then
            MsgBox("�������� ������ �� WEB ����������� �������.", MsgBoxStyle.Information, "��������!")
        Else
            MsgBox("�� ����� �������� ������ �� WEB ���� ������.", MsgBoxStyle.Information, "��������!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� �� ���������� � ����������� �� ������� - ������ � ������������ ������� � ��������� �� WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If AvailabilityUploadToCatalog_WEB(1) = True Then
            MsgBox("�������� ������ �� WEB ����������� �������.", MsgBoxStyle.Information, "��������!")
        Else
            MsgBox("�� ����� �������� ������ �� WEB ���� ������.", MsgBoxStyle.Information, "��������!")
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button75.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� �� ���������� ��� ����������� �������� � ��� ������� � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadInfoToSaintGobain = New UploadInfoToSaintGobain
        MyUploadInfoToSaintGobain.ShowDialog()
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� �������,
        '// ������� ������� ����������������� � ������ ������ ���������� � ����� ��.
        '// �������� ����� ������������� ���� ������ ���������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDeletePictures = New DeletePictures
        MyDeletePictures.ShowDialog()
    End Sub

   
    Private Sub Button77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button77.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� �� �������� �������
        '// � ��������� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadFilesToCatalog = New UploadFilesToCatalog
        MyUploadFilesToCatalog.ShowDialog()
    End Sub

    Private Sub Button76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button76.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � �� �������� �������
        '// �������� �������������� ����������������� � .jpg
        '// �������� ����� ������������� ���� ������ ���������� (��� ���� � ����, ��� �������� ������������� �����?)
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadFilesToDB = New UploadFilesToDB
        MyUploadFilesToDB.ShowDialog()
    End Sub

    Private Sub Button79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button79.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� � �� �������� � �������� �������
        '// ���������� �� WEB
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyTransferNamesDescrToDB = New TransferNamesDescrToDB
        MyTransferNamesDescrToDB.ShowDialog()
    End Sub

    Private Sub Button80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button80.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ���������� �������� � ������ Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyMatchPictAndScalaCode = New MatchPictAndScalaCode
        MyMatchPictAndScalaCode.ShowDialog()
    End Sub

    Private Sub Button82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button82.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �������� �������� �� ��
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyDeletePictureFromDB = New DeletePictureFromDB
        MyDeletePictureFromDB.ShowDialog()
    End Sub

    Private Sub Button81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button81.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� �������� ����� �������� � ��
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyLoadOnePictToDB = New LoadOnePictToDB
        MyLoadOnePictToDB.ShowDialog()
    End Sub

    Private Sub Button83_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button83.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � �������
        '// ������ � ������� ������� ��������
        '// ����������� �������� (�������� - ��� ������ ����������) 
        '// ��������, �������� - � Excel � ��� �� ��������
        '////////////////////////////////////////////////////////////////////////////////

        MyDownloadInfoFromSE = New DownloadInfoFromSE
        MyDownloadInfoFromSE.ShowDialog()
    End Sub

    Private Sub Button85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button85.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���� Magento ������ ����� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadDataToMagento = New UploadDataToMagento
        MyUploadDataToMagento.MyMode = 1
        MyUploadDataToMagento.ShowDialog()
    End Sub

    Private Sub Button84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button84.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadDataToMagento = New UploadDataToMagento
        MyUploadDataToMagento.MyMode = 0
        MyUploadDataToMagento.ShowDialog()
    End Sub

    Private Sub Button86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button86.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� � ����������� �� ������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadAvailabilityToMagento = New UploadAvailabilityToMagento
        MyUploadAvailabilityToMagento.ShowDialog()
    End Sub

    Private Sub Button87_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button87.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyUploadPicturesToMagento = New UploadPicturesToMagento
        MyUploadPicturesToMagento.ShowDialog()
    End Sub

    Private Sub Button88_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button88.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���� ������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub Button89_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button89.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����������������� ������������ �� Scala 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCASH_FullUpload = New CASH_FullUpload
        MyCASH_FullUpload.ShowDialog()

    End Sub

    Private Sub Button90_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button90.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCASH_CustomUpload = New CASH_CustomUpload
        MyCASH_CustomUpload.ShowDialog()
    End Sub

    Private Sub Button91_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button91.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� ���������� ����  
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
            MsgBox("����� " & Trim(TextBox5.Text) & " �� ������. ", MsgBoxStyle.Critical, "��������!")
        Else
            '----- ����������� ������ ������
            For i As Integer = 0 To DataGridView7.Rows.Count - 1
                If DataGridView7.Item(0, i).Value = Trim(Declarations.MyRec.Fields("GroupCode").Value) Then
                    DataGridView7.CurrentCell = DataGridView7.Item(1, i)
                    Exit For
                End If
            Next

            '----- ����������� ��������� ������ ��� ����������� ������
            If Trim(Declarations.MyRec.Fields("GroupCode").Value).Equals("") Then
                '-----����� �� � ��������� - ������� ��� � ������� ��� ���������
                For i As Integer = 0 To DataGridView9.Rows.Count - 1
                    If DataGridView9.Item(1, i).Value = Trim(TextBox5.Text) Then
                        DataGridView9.CurrentCell = DataGridView9.Item(1, i)
                        Exit For
                    End If
                Next
            Else
                '-----����� � ���������
                '-----������� ������� ��������� � ����������
                For i As Integer = 0 To DataGridView8.Rows.Count - 1
                    If DataGridView8.Item(0, i).Value = Trim(Declarations.MyRec.Fields("SubgroupID").Value) Then
                        DataGridView8.CurrentCell = DataGridView8.Item(1, i)
                        Exit For
                    End If
                Next

                '-----������� ����� � ������� � ����������
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
        '// �������� � Excel ���������� � �����, ������, ������ � ���� ������  
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
        '// �������� �� Excel ���������� � �����, ������, ������ � ���� ������  
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            LoadItemDimFromLO()
        Else
            LoadItemDimFromExcel()
        End If
        MsgBox("�������� ������ �� ��������� ���������", MsgBoxStyle.Information, "��������!")
        Cursor = Cursors.Default
    End Sub

    Private Sub Button94_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button94.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � �������
        '// ������ � ������� ABB
        '// ����������� �������� (�������� - ��� ������ ����������) 
        '// ��������, �������� - � Excel � ��� �� ��������
        '////////////////////////////////////////////////////////////////////////////////

        MyDownloadInfoFromABB = New DownloadInfoFromABB
        MyDownloadInfoFromABB.ShowDialog()
    End Sub

End Class


