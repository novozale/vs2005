Imports System.IO

Public Class UploadFilesToDB

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� ����� ��� �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("������� � ���������� ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Function CheckDataMF() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� ����� ��� ��������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
            MsgBox("������� � ���������� ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckDataMF = False
            TextBox2.Select()
            Exit Function
        End If

        CheckDataMF = True
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� �� ���������� �������� - ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyParam As Integer

        If CheckData() = True Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False

            Select Case ComboBox2.SelectedItem
                Case "�� ��������������"
                    MyParam = 0
                Case "������������ ����������� � ����� Scala"
                    MyParam = 1
                Case "������������ ���"
                    MyParam = 2
                Case Else
                    MyParam = 0
            End Select

            UploadMyPictures(Trim(TextBox1.Text), MyParam, Trim(ComboBox1.SelectedValue))
            MsgBox("�������� �������� ���������.", MsgBoxStyle.Information, "��������!")
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
        End If
    End Sub

    Private Sub UploadMyPictures(ByVal MyCatalog As String, ByVal MyParam As Integer, ByVal MySupplier As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �������� �� ���������� �������� - ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim Folder As Directory
        Dim Files() As String
        Dim i As Integer

        Files = Folder.GetFiles(MyCatalog, "*.jpg")
        If Files.Length = 0 Then
            MsgBox("� ��������� �������� ��� �� ������ ����� � ����������� jpg.", MsgBoxStyle.Critical, "��������!")
        Else
            Label3.Text = CStr(Files.Length)
            For i = 0 To Files.Length - 1
                UploadOnePictureToDB(Files(i), FileIO.FileSystem.GetName(Files(i)).Substring(0, Len(FileIO.FileSystem.GetName(Files(i))) - 4), MyParam, MySupplier)
                Label2.Text = CStr(i)
                Application.DoEvents()
            Next
        End If
    End Sub

    Public Function MkTh(ByVal MyImg As Byte(), ByVal MyWidth As Integer, ByVal MyHeight As Integer) As Byte()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ����������� ����� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim ms As MemoryStream = New MemoryStream()
        Dim thumbnail As Image = Image.FromStream(New MemoryStream(MyImg)).GetThumbnailImage(MyWidth, MyHeight, Nothing, New IntPtr())

        thumbnail.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        MkTh = ms.ToArray()

    End Function

    Private Sub UploadFilesToDB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������ �� alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub UploadFilesToDB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '��� �����������
        Dim MyDs1 As New DataSet
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    '��� ��������������
        Dim MyDs2 As New DataSet

        InitMyConn(False)
        '---����������
        MySQLStr = "SELECT '---' AS Code, ' ���' AS Name "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT DISTINCT SC010300.SC01058 AS Code, SC010300.SC01058 + ' ' + PL010300.PL01002 AS Name "
        MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "ORDER BY Code "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox1.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---�������������
        MySQLStr = "SELECT '---' AS Code, ' ���' AS ManName "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT Convert(nvarchar(30), ID) AS Code, LTRIM(RTRIM(LTRIM(RTRIM(Name)) + ' ' + LTRIM(RTRIM(Address)))) AS ManName "
        MySQLStr = MySQLStr & "FROM tbl_Manufacturers "
        MySQLStr = MySQLStr & "ORDER BY ManName "
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.SelectCommand.CommandTimeout = 600
            MyAdapter2.Fill(MyDs2)
            ComboBox3.DisplayMember = "ManName" '��� �� ��� ����� ������������
            ComboBox3.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox3.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ComboBox2.SelectedItem = "�� ��������������"
        ComboBox4.SelectedItem = "�� ��������������"
    End Sub

    Private Sub UploadOnePictureToDB(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal MyParam As Integer, ByVal MySupplier As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ����� �������� � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyTotalQTY As Integer
        Dim MyNotMatchedQTY As Integer
        Dim MyMatchedQTY As Integer

        If MySupplier = "---" Then          '---��� ���� �����������
            '-------------��������� ������� �������� � ��, ���-�� ������������� ����� ����������, ��������� ��������
            MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM (SELECT SC01060, COUNT(SC01060) AS TotalQTY "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS NotMatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
            MySQLStr = MySQLStr & "WHERE (SupplierItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "AND (ScalaItemCode IS NULL) "
            MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 "
            MySQLStr = MySQLStr & "WHERE (SupplierItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "AND (NOT (ScalaItemCode IS NULL)) "
            MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '---��� ������� ������ - �� � �� ����� ������ �������
                trycloseMyRec()
                Exit Sub
            Else
                MyTotalQTY = Declarations.MyRec.Fields("TotalQTY").Value
                MyNotMatchedQTY = Declarations.MyRec.Fields("NotMatchedQTY").Value
                MyMatchedQTY = Declarations.MyRec.Fields("MatchedQTY").Value
                trycloseMyRec()
            End If
        Else                                '---��� ���������� ����������
            '-------------��������� ������� �������� � ��, ���-�� ������������� ����� ����������, ��������� ��������
            MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM (SELECT SC01060, COUNT(SC01060) AS TotalQTY "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "AND (SC01058 = N'" & Trim(MySupplier) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC01060) AS View_1 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS NotMatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
            MySQLStr = MySQLStr & "WHERE (SupplierItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "AND (ScalaItemCode IS NULL) "
            MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures_1.SupplierItemCode, COUNT(tbl_WEB_Pictures_1.SupplierItemCode) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON tbl_WEB_Pictures_1.ScalaItemCode = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures_1.SupplierItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & Trim(MySupplier) & "') "
            MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures_1.SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '---��� ������� ������ - �� � �� ����� ������ �������
                trycloseMyRec()
                Exit Sub
            Else
                MyTotalQTY = Declarations.MyRec.Fields("TotalQTY").Value
                MyNotMatchedQTY = Declarations.MyRec.Fields("NotMatchedQTY").Value
                MyMatchedQTY = Declarations.MyRec.Fields("MatchedQTY").Value
                trycloseMyRec()
                '---��� ��� ����������� �������� (�� �� ���������, � ������ ���������� ��� ���������), ����� ���� ������,
                '---��� ����� ���������� ������� � ����� ���������� �� ������ ����������, ��:
                'If MyNotMatchedQTY > (MyTotalQTY - MyMatchedQTY) Then
                '     MyNotMatchedQTY = (MyTotalQTY - MyMatchedQTY)
                ' End If
            End If
        End If

        '-------����������, ����� �������� ����� ���������-----------------------------------------------
        If MyParam = 0 Then             '---������ �� ��������������
            If MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then   '---��� ���� ����������� �������� - ���-�� ������� � ����� ����� ���������� ������, ��� �������� � ��
                WritePictureToDB(MyPicturePath, MyPictureName)
            End If
        ElseIf MyParam = 1 Then         '---������������ ������ ����������� � ����� Scala
            If MyNotMatchedQTY > 0 Then     '---���� ��� ������������
                UpdatePictureInDB(MyPicturePath, MyPictureName, 0, MySupplier)
            ElseIf MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then '---�������������� ������ - �������, ���� �����
                WritePictureToDB(MyPicturePath, MyPictureName)
            End If
        ElseIf MyParam = 2 Then         '---������������ ���
            If (MyNotMatchedQTY + MyMatchedQTY) > 0 Then    '---���� ��� ������������
                UpdatePictureInDB(MyPicturePath, MyPictureName, 1, MySupplier)
            ElseIf MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then '---�������������� ������ - �������, ���� �����
                WritePictureToDB(MyPicturePath, MyPictureName)
            End If
        End If

    End Sub

    Private Sub WritePictureToDB(ByVal MyPicturePath As String, ByVal MyPictureName As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������ ����� �������� � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim connection As SqlClient.SqlConnection

        Try
            Connection = New SqlClient.SqlConnection(Declarations.MyNETConnStr)

            Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand("INSERT INTO tbl_WEB_Pictures " _
                + "(ID, ScalaItemCode, SupplierItemCode, PictureSmall, PictureMedium, Picture) " _
                + "VALUES(NEWID(), NULL, @SupplierItemCode, @PictureSmallBytes, @PictureMediumBytes, @PictureBytes) ", connection)

            Dim par As SqlClient.SqlParameter = New SqlClient.SqlParameter("@SupplierItemCode", SqlDbType.NVarChar)
            par.Value = MyPictureName.ToString
            par.Direction = ParameterDirection.Input
            cmd.Parameters.Add(par)

            par = New SqlClient.SqlParameter("@PictureBytes", SqlDbType.Image)
            par.Direction = ParameterDirection.Input
            Dim fStream As FileStream = New FileStream(MyPicturePath, FileMode.Open, FileAccess.Read)
            Dim lBytes As Long = fStream.Length
            If (lBytes > 0) Then
                Dim imageBytes(lBytes - 1) As Byte
                fStream.Read(imageBytes, 0, lBytes)
                fStream.Close()
                par.Value = imageBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureMediumBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageMediumBytes(10000) As Byte
                imageMediumBytes = MkTh(imageBytes, 100, 100)
                par.Value = imageMediumBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureSmallBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageSmallBytes(1225) As Byte
                imageSmallBytes = MkTh(imageBytes, 35, 35)
                par.Value = imageSmallBytes
                cmd.Parameters.Add(par)

                connection.Open()
                cmd.ExecuteNonQuery()
            Else
                connection.Dispose()
                MsgBox("������ ����� " & MyPicturePath & " ����� 0. ����� ���� �� ����� ��������.")
            End If
            connection.Dispose()
        Catch ex As Exception
            Try
                Connection.Dispose()
            Catch
            End Try
            MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
        End Try
    End Sub

    Private Sub UpdatePictureInDB(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal MyParam As Integer, ByVal MySupplier As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ����� �������� � ��
        '// MyParam = 0 - ��������� ������ ����������� ��������
        '// MyParam = 1 - ��������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim connection As SqlClient.SqlConnection
        Dim MySQLStr As String

        Try
            connection = New SqlClient.SqlConnection(Declarations.MyNETConnStr)

            If MyParam = 0 Then
                MySQLStr = "Update tbl_WEB_Pictures "
                MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IS NULL) "
            ElseIf MyParam = 1 Then
                If MySupplier = "---" Then          '---��� ���� �����������
                    MySQLStr = "Update tbl_WEB_Pictures "
                    MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) "
                Else                                '---��� ���������� ����������
                    MySQLStr = "UPDATE tbl_WEB_Pictures "
                    MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IS NULL) OR "
                    MySQLStr = MySQLStr & "(SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IN "
                    MySQLStr = MySQLStr & "(SELECT SC01001 "
                    MySQLStr = MySQLStr & "FROM SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01058 = @SupplierCode))) "
                End If
            End If

            Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand(MySQLStr, connection)

            Dim par As SqlClient.SqlParameter = New SqlClient.SqlParameter("@SupplierItemCode", SqlDbType.NVarChar)
            par.Value = MyPictureName.ToString
            par.Direction = ParameterDirection.Input
            cmd.Parameters.Add(par)

            If MyParam = 1 And MySupplier <> "---" Then     '---�������������� �������� - ��� ����������
                par = New SqlClient.SqlParameter("@SupplierCode", SqlDbType.NVarChar)
                par.Direction = ParameterDirection.Input
                par.Value = MySupplier
                cmd.Parameters.Add(par)
            End If

            par = New SqlClient.SqlParameter("@PictureBytes", SqlDbType.Image)
            par.Direction = ParameterDirection.Input
            Dim fStream As FileStream = New FileStream(MyPicturePath, FileMode.Open, FileAccess.Read)
            Dim lBytes As Long = fStream.Length
            If (lBytes > 0) Then
                Dim imageBytes(lBytes - 1) As Byte
                fStream.Read(imageBytes, 0, lBytes)
                fStream.Close()
                par.Value = imageBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureMediumBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageMediumBytes(10000) As Byte
                imageMediumBytes = MkTh(imageBytes, 100, 100)
                par.Value = imageMediumBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureSmallBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageSmallBytes(625) As Byte
                imageSmallBytes = MkTh(imageBytes, 25, 25)
                par.Value = imageSmallBytes
                cmd.Parameters.Add(par)

                connection.Open()
                cmd.ExecuteNonQuery()
            Else
                connection.Dispose()
                MsgBox("������ ����� " & MyPicturePath & " ����� 0. ����� ���� �� ����� ��������.")
            End If
            connection.Dispose()
        Catch ex As Exception
            Try
                connection.Dispose()
            Catch
            End Try
            MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �������� � ���������� ��� �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �������� � ���������� ��� ��������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
        Else
            TextBox2.Text = MyCatalog
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� �� ���������� �������� - �������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyParam As Integer

        If CheckDataMF() = True Then
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False

            Select Case ComboBox4.SelectedItem
                Case "�� ��������������"
                    MyParam = 0
                Case "������������ ����������� � ����� Scala"
                    MyParam = 1
                Case "������������ ���"
                    MyParam = 2
                Case Else
                    MyParam = 0
            End Select

            UploadMyPicturesMF(Trim(TextBox2.Text), MyParam, Trim(ComboBox3.SelectedValue))
            MsgBox("�������� �������� ���������.", MsgBoxStyle.Information, "��������!")
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub UploadMyPicturesMF(ByVal MyCatalog As String, ByVal MyParam As Integer, ByVal MyManufacturer As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �������� �� ���������� �������� - �������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim Folder As Directory
        Dim Files() As String
        Dim i As Integer

        Files = Folder.GetFiles(MyCatalog, "*.jpg")
        If Files.Length = 0 Then
            MsgBox("� ��������� �������� ��� �� ������ ����� � ����������� jpg.", MsgBoxStyle.Critical, "��������!")
        Else
            Label10.Text = CStr(Files.Length)
            For i = 0 To Files.Length - 1
                UploadOnePictureToDBMF(Files(i), FileIO.FileSystem.GetName(Files(i)).Substring(0, Len(FileIO.FileSystem.GetName(Files(i))) - 4), MyParam, MyManufacturer)
                Label11.Text = CStr(i)
                Application.DoEvents()
            Next
        End If
    End Sub

    Private Sub UploadOnePictureToDBMF(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal MyParam As Integer, ByVal MyManufacturer As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �������� � �� - �������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyTotalQTY As Integer
        Dim MyNotMatchedQTY As Integer
        Dim MyMatchedQTY As Integer
        Dim MySuppPictureName As String

        If MyManufacturer = "---" Then          '---��� ���� ��������������
            '-------------��������� ������� �������� � ��, ���-�� ������������� ����� �������������, ��������� ��������
            'MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            'MySQLStr = MySQLStr & "FROM (SELECT SC010300.SC01060, COUNT(SC010300.SC01060) AS TotalQTY "
            'MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "
            'MySQLStr = MySQLStr & "WHERE (SC010300.SC01060 = N'" & Trim(MyPictureName) & "') "
            'MySQLStr = MySQLStr & "GROUP BY SC010300.SC01060) AS View_1 LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS NotMatchedQTY "
            'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
            'MySQLStr = MySQLStr & "WHERE (SupplierItemCode = N'" & Trim(MyPictureName) & "') AND (ScalaItemCode IS NULL) "
            'MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures_1.SupplierItemCode, COUNT(tbl_WEB_Pictures_1.SupplierItemCode) AS MatchedQTY "
            'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 INNER JOIN "
            'MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEB_Pictures_1.ScalaItemCode = SC010300_1.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_1 ON SC010300_1.SC01001 = tbl_ItemCard0300_1.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_1 ON tbl_ItemCard0300_1.Manufacturer = tbl_Manufacturers_1.ID "
            'MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures_1.SupplierItemCode = N'" & Trim(MyPictureName) & "') "
            'MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures_1.SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "

            MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM (SELECT SC010300.SC01060, COUNT(SC010300.SC01060) AS TotalQTY, tbl_ItemCard0300.ManufacturerItemCode "
            MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_ItemCard0300.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC010300.SC01060, tbl_ItemCard0300.ManufacturerItemCode) AS View_1 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures.SupplierItemCode, COUNT(tbl_WEB_Pictures.SupplierItemCode) AS NotMatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT SC010300_2.SC01060, COUNT(SC010300_2.SC01060) AS TotalQTY, tbl_ItemCard0300_2.ManufacturerItemCode "
            MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_2 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_2 ON SC010300_2.SC01001 = tbl_ItemCard0300_2.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_2 ON tbl_ItemCard0300_2.Manufacturer = tbl_Manufacturers_2.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_ItemCard0300_2.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC010300_2.SC01060, tbl_ItemCard0300_2.ManufacturerItemCode) AS View_1_1 ON "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures.SupplierItemCode = View_1_1.SC01060 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode Is NULL) "
            MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures.SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures_1.SupplierItemCode, COUNT(tbl_WEB_Pictures_1.SupplierItemCode) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEB_Pictures_1.ScalaItemCode = SC010300_1.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_1 ON SC010300_1.SC01001 = tbl_ItemCard0300_1.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_1 ON tbl_ItemCard0300_1.Manufacturer = tbl_Manufacturers_1.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_ItemCard0300_1.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures_1.SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '---��� ������� ������ - �� � �� ����� ������ �������
                trycloseMyRec()
                Exit Sub
            Else
                MyTotalQTY = Declarations.MyRec.Fields("TotalQTY").Value
                MyNotMatchedQTY = Declarations.MyRec.Fields("NotMatchedQTY").Value
                MyMatchedQTY = Declarations.MyRec.Fields("MatchedQTY").Value
                MySuppPictureName = Declarations.MyRec.Fields("SC01060").Value
                trycloseMyRec()
            End If
        Else                                '---��� ���������� �������������
            '-------------��������� ������� �������� � ��, ���-�� ������������� ����� �������������, ��������� ��������
            'MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            'MySQLStr = MySQLStr & "FROM (SELECT SC010300.SC01060, COUNT(SC010300.SC01060) AS TotalQTY "
            'MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "
            'MySQLStr = MySQLStr & "WHERE (SC010300.SC01060 = N'" & Trim(MyPictureName) & "') AND (tbl_Manufacturers.ID = " & Trim(MyManufacturer) & ") "
            'MySQLStr = MySQLStr & "GROUP BY SC010300.SC01060) AS View_1 LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "(SELECT SupplierItemCode, COUNT(SupplierItemCode) AS NotMatchedQTY "
            'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
            'MySQLStr = MySQLStr & "WHERE (SupplierItemCode = N'" & Trim(MyPictureName) & "') AND (ScalaItemCode IS NULL) "
            'MySQLStr = MySQLStr & "GROUP BY SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            'MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures_1.SupplierItemCode, COUNT(tbl_WEB_Pictures_1.SupplierItemCode) AS MatchedQTY "
            'MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 INNER JOIN "
            'MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEB_Pictures_1.ScalaItemCode = SC010300_1.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_1 ON SC010300_1.SC01001 = tbl_ItemCard0300_1.SC01001 INNER JOIN "
            'MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_1 ON tbl_ItemCard0300_1.Manufacturer = tbl_Manufacturers_1.ID "
            'MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures_1.SupplierItemCode = N'" & Trim(MyPictureName) & "') AND (tbl_Manufacturers_1.ID = " & Trim(MyManufacturer) & ") "
            'MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures_1.SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "

            MySQLStr = "SELECT View_1.SC01060, View_1.TotalQTY, ISNULL(View_2.NotMatchedQTY, 0) AS NotMatchedQTY, ISNULL(View_3.MatchedQTY, 0) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM (SELECT SC010300.SC01060, COUNT(SC010300.SC01060) AS TotalQTY, tbl_ItemCard0300.ManufacturerItemCode "
            MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_Manufacturers.ID = " & Trim(MyManufacturer) & ") AND (tbl_ItemCard0300.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC010300.SC01060, tbl_ItemCard0300.ManufacturerItemCode) AS View_1 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures.SupplierItemCode, COUNT(tbl_WEB_Pictures.SupplierItemCode) AS NotMatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT SC010300_2.SC01060, COUNT(SC010300_2.SC01060) AS TotalQTY, tbl_ItemCard0300_2.ManufacturerItemCode "
            MySQLStr = MySQLStr & "FROM SC010300 AS SC010300_2 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_2 ON SC010300_2.SC01001 = tbl_ItemCard0300_2.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_2 ON tbl_ItemCard0300_2.Manufacturer = tbl_Manufacturers_2.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_Manufacturers_2.ID = " & Trim(MyManufacturer) & ") AND (tbl_ItemCard0300_2.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY SC010300_2.SC01060, tbl_ItemCard0300_2.ManufacturerItemCode) AS View_1_1 ON "
            MySQLStr = MySQLStr & "tbl_WEB_Pictures.SupplierItemCode = View_1_1.SC01060 "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode Is NULL) "
            MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures.SupplierItemCode) AS View_2 ON View_1.SC01060 = View_2.SupplierItemCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT tbl_WEB_Pictures_1.SupplierItemCode, COUNT(tbl_WEB_Pictures_1.SupplierItemCode) AS MatchedQTY "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures AS tbl_WEB_Pictures_1 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 AS SC010300_1 ON tbl_WEB_Pictures_1.ScalaItemCode = SC010300_1.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_ItemCard0300 AS tbl_ItemCard0300_1 ON SC010300_1.SC01001 = tbl_ItemCard0300_1.SC01001 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_Manufacturers AS tbl_Manufacturers_1 ON tbl_ItemCard0300_1.Manufacturer = tbl_Manufacturers_1.ID "
            MySQLStr = MySQLStr & "WHERE (tbl_Manufacturers_1.ID = " & Trim(MyManufacturer) & ") AND (tbl_ItemCard0300_1.ManufacturerItemCode = N'" & Trim(MyPictureName) & "') "
            MySQLStr = MySQLStr & "GROUP BY tbl_WEB_Pictures_1.SupplierItemCode) AS View_3 ON View_1.SC01060 = View_3.SupplierItemCode "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                '---��� ������� ������ - �� � �� ����� ������ �������
                trycloseMyRec()
                Exit Sub
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    MyTotalQTY = Declarations.MyRec.Fields("TotalQTY").Value
                    MyNotMatchedQTY = Declarations.MyRec.Fields("NotMatchedQTY").Value
                    MyMatchedQTY = Declarations.MyRec.Fields("MatchedQTY").Value
                    MySuppPictureName = Declarations.MyRec.Fields("SC01060").Value
                    PrepareWritePictureToDBMF(MyPicturePath, MySuppPictureName, MyTotalQTY, MyNotMatchedQTY, _
                        MyMatchedQTY, MyParam, MyManufacturer)
 
                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()
            End If
        End If
    End Sub

    Private Sub PrepareWritePictureToDBMF(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal MyTotalQTY As Integer, _
    ByVal MyNotMatchedQTY As Integer, ByVal MyMatchedQTY As Integer, ByVal MyParam As Integer, ByVal MyManufacturer As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� �������� �������� � �� - �������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '-------����������, ����� �������� ����� ���������-----------------------------------------------
        If MyParam = 0 Then             '---������ �� ��������������
            If MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then   '---��� ���� ����������� �������� - ���-�� ������� � ����� ����� ������������� ������, ��� �������� � ��
                WritePictureToDBMF(MyPicturePath, MyPictureName)
            End If
        ElseIf MyParam = 1 Then         '---������������ ������ ����������� � ����� Scala
            If MyNotMatchedQTY > 0 Then     '---���� ��� ������������
                UpdatePictureInDBMF(MyPicturePath, MyPictureName, 0, MyManufacturer)
            ElseIf MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then '---�������������� ������ - �������, ���� �����
                WritePictureToDBMF(MyPicturePath, MyPictureName)
            End If
        ElseIf MyParam = 2 Then         '---������������ ���
            If (MyNotMatchedQTY + MyMatchedQTY) > 0 Then    '---���� ��� ������������
                UpdatePictureInDBMF(MyPicturePath, MyPictureName, 1, MyManufacturer)
            ElseIf MyTotalQTY - (MyNotMatchedQTY + MyMatchedQTY) > 0 Then '---�������������� ������ - �������, ���� �����
                WritePictureToDBMF(MyPicturePath, MyPictureName)
            End If
        End If
    End Sub

    Private Sub WritePictureToDBMF(ByVal MyPicturePath As String, ByVal MyPictureName As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������ ����� �������� � �� - �� ��������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim connection As SqlClient.SqlConnection

        Try
            connection = New SqlClient.SqlConnection(Declarations.MyNETConnStr)

            Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand("INSERT INTO tbl_WEB_Pictures " _
                + "(ID, ScalaItemCode, SupplierItemCode, PictureSmall, PictureMedium, Picture) " _
                + "VALUES(NEWID(), NULL, @SupplierItemCode, @PictureSmallBytes, @PictureMediumBytes, @PictureBytes) ", connection)

            Dim par As SqlClient.SqlParameter = New SqlClient.SqlParameter("@SupplierItemCode", SqlDbType.NVarChar)
            par.Value = MyPictureName.ToString
            par.Direction = ParameterDirection.Input
            cmd.Parameters.Add(par)

            par = New SqlClient.SqlParameter("@PictureBytes", SqlDbType.Image)
            par.Direction = ParameterDirection.Input
            Dim fStream As FileStream = New FileStream(MyPicturePath, FileMode.Open, FileAccess.Read)
            Dim lBytes As Long = fStream.Length
            If (lBytes > 0) Then
                Dim imageBytes(lBytes - 1) As Byte
                fStream.Read(imageBytes, 0, lBytes)
                fStream.Close()
                par.Value = imageBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureMediumBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageMediumBytes(10000) As Byte
                imageMediumBytes = MkTh(imageBytes, 100, 100)
                par.Value = imageMediumBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureSmallBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageSmallBytes(1225) As Byte
                imageSmallBytes = MkTh(imageBytes, 35, 35)
                par.Value = imageSmallBytes
                cmd.Parameters.Add(par)

                connection.Open()
                cmd.ExecuteNonQuery()
            Else
                connection.Dispose()
                MsgBox("������ ����� " & MyPicturePath & " ����� 0. ����� ���� �� ����� ��������.")
            End If
            connection.Dispose()
        Catch ex As Exception
            Try
                connection.Dispose()
            Catch
            End Try
            MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
        End Try
    End Sub

    Private Sub UpdatePictureInDBMF(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal MyParam As Integer, ByVal MyManufacturer As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ����� �������� � ��
        '// MyParam = 0 - ��������� ������ ����������� ��������
        '// MyParam = 1 - ��������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim connection As SqlClient.SqlConnection
        Dim MySQLStr As String

        Try
            connection = New SqlClient.SqlConnection(Declarations.MyNETConnStr)

            If MyParam = 0 Then
                MySQLStr = "Update tbl_WEB_Pictures "
                MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IS NULL) "
            ElseIf MyParam = 1 Then
                If MyManufacturer = "---" Then          '---��� ���� ��������������
                    'MySQLStr = "Update tbl_WEB_Pictures "
                    'MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    'MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) "
                    MySQLStr = "Update tbl_WEB_Pictures "
                    MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) "
                    MySQLStr = MySQLStr & "AND (ScalaItemCode IS NULL) "
                Else                                '---��� ���������� �������������
                    'MySQLStr = "UPDATE tbl_WEB_Pictures "
                    'MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    'MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IS NULL) OR "
                    'MySQLStr = MySQLStr & "(SupplierItemCode = @SupplierItemCode) AND (ScalaItemCode IN "
                    'MySQLStr = MySQLStr & "(SELECT SC01001 "
                    'MySQLStr = MySQLStr & "FROM SC010300 "
                    'MySQLStr = MySQLStr & "WHERE (SC01058 = @SupplierCode))) "
                    MySQLStr = "UPDATE tbl_WEB_Pictures "
                    MySQLStr = MySQLStr & "SET PictureSmall = @PictureSmallBytes, PictureMedium = @PictureMediumBytes, Picture = @PictureBytes "
                    MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ItemCard0300 ON tbl_WEB_Pictures.ScalaItemCode = tbl_ItemCard0300.SC01001 "
                    MySQLStr = MySQLStr & "WHERE (SupplierItemCode = @SupplierItemCode) "
                    MySQLStr = MySQLStr & "AND (tbl_ItemCard0300.Manufacturer = @ManufacturerCode) "
                End If
            End If

            Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand(MySQLStr, connection)

            Dim par As SqlClient.SqlParameter = New SqlClient.SqlParameter("@SupplierItemCode", SqlDbType.NVarChar)
            par.Value = MyPictureName.ToString
            par.Direction = ParameterDirection.Input
            cmd.Parameters.Add(par)

            If MyParam = 1 And MyManufacturer <> "---" Then     '---�������������� �������� - ��� �������������
                par = New SqlClient.SqlParameter("@ManufacturerCode", SqlDbType.NVarChar)
                par.Direction = ParameterDirection.Input
                par.Value = MyManufacturer
                cmd.Parameters.Add(par)
            End If

            par = New SqlClient.SqlParameter("@PictureBytes", SqlDbType.Image)
            par.Direction = ParameterDirection.Input
            Dim fStream As FileStream = New FileStream(MyPicturePath, FileMode.Open, FileAccess.Read)
            Dim lBytes As Long = fStream.Length
            If (lBytes > 0) Then
                Dim imageBytes(lBytes - 1) As Byte
                fStream.Read(imageBytes, 0, lBytes)
                fStream.Close()
                par.Value = imageBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureMediumBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageMediumBytes(10000) As Byte
                imageMediumBytes = MkTh(imageBytes, 100, 100)
                par.Value = imageMediumBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureSmallBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageSmallBytes(625) As Byte
                imageSmallBytes = MkTh(imageBytes, 25, 25)
                par.Value = imageSmallBytes
                cmd.Parameters.Add(par)

                connection.Open()
                cmd.ExecuteNonQuery()
            Else
                connection.Dispose()
                MsgBox("������ ����� " & MyPicturePath & " ����� 0. ����� ���� �� ����� ��������.")
            End If
            connection.Dispose()
        Catch ex As Exception
            Try
                connection.Dispose()
            Catch
            End Try
            MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
        End Try
    End Sub
End Class