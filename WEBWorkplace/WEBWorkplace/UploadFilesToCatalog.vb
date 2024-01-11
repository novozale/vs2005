Imports System.IO

Public Class UploadFilesToCatalog

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("������� ��� �������� ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub UploadFilesToCatalog_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������ �� alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub UploadFilesToCatalog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ������� ������� � �����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ����� �������
        Dim MyDs As New DataSet
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    '��� �����������
        Dim MyDs1 As New DataSet

        InitMyConn(False)
        '---������ �������
        MySQLStr = "SELECT '---' AS Code, '���' AS Name "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT SY24002 AS Code, SY24002 + '  ' + SY24003 AS Name "
        MySQLStr = MySQLStr & "FROM SY240300 "
        MySQLStr = MySQLStr & "WHERE (SY24001 = N'IB') AND (SY24002 <> N'') "
        MySQLStr = MySQLStr & "ORDER BY Code "
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

        '---����������
        MySQLStr = "SELECT '---' AS Code, '���' AS Name "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT DISTINCT SC010300.SC01058 AS Code, SC010300.SC01058 + ' ' + PL010300.PL01002 AS Name "
        MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 "
        MySQLStr = MySQLStr & "ORDER BY Code "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "Code"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ComboBox3.SelectedItem = "���� Scala"
        ComboBox4.SelectedItem = "��� ������"
        ComboBox5.SelectedItem = "���������� -> ������ �������"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �������� � ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������� � �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCodes As Integer
        Dim MyRange As Integer
        Dim MyGroup As Integer
        Dim MySupplCatCreated As Integer
        Dim MyGroupCatCreated As Integer

        If CheckData() = True Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False

            Select Case ComboBox3.SelectedItem
                Case "���� Scala"
                    MyCodes = 0
                Case "���� ������ ����������"
                    MyCodes = 1
                Case Else
                    MyCodes = 0
            End Select

            Select Case ComboBox4.SelectedItem
                Case "��� ������"
                    MyRange = 0
                Case "������ ������ �� ����� �����"
                    MyRange = 1
                Case Else
                    MyRange = 0
            End Select

            Select Case ComboBox5.SelectedItem
                Case "���������� -> ������ �������"
                    MyGroup = 0
                Case "������ ������� -> ����������"
                    MyGroup = 1
                Case Else
                    MyGroup = 0
            End Select

            If CheckBox1.Checked = False Then
                MySupplCatCreated = 0
            Else
                MySupplCatCreated = 1
            End If

            If CheckBox2.Checked = False Then
                MyGroupCatCreated = 0
            Else
                MyGroupCatCreated = 1
            End If

            DownloadMyPictures(Trim(TextBox1.Text), ComboBox1.SelectedValue, ComboBox2.SelectedValue, MyCodes, MyRange, MyGroup, MySupplCatCreated, MyGroupCatCreated)
            MsgBox("�������� �������� ���������.", MsgBoxStyle.Information, "��������!")

            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
        End If
    End Sub

    Private Sub DownloadMyPictures(ByVal MyCatalog As String, ByVal MyGroup As String, ByVal MySupplier As String, ByVal MyCodes As Integer, ByVal MyRange As Integer, _
    ByVal MyGroupDir As Integer, ByVal MySupplCatCreated As Integer, ByVal MyGroupCatCreated As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �������� � �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        If MyRange = 0 Then         '---��� ������ �� ���� ��� �����������
            If MyCodes = 0 Then         '---��������� �������� � ������ Scala
                MySQLStr = "SELECT tbl_WEB_Pictures.ScalaItemCode AS ItemCode, tbl_WEB_Pictures.Picture, Ltrim(Rtrim(ISNULL(View_1.SY24002, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(View_1.SY24003, "
                MySQLStr = MySQLStr & "N'����������'))) AS ProductGroup, Ltrim(Rtrim(ISNULL(PL010300.PL01001, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(PL010300.PL01002, N'����������'))) AS Supplier "
                MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SY24002, SY24003 "
                MySQLStr = MySQLStr & "FROM SY240300 "
                MySQLStr = MySQLStr & "WHERE (SY24001 = N'IB') AND (SY24002 <> N'')) AS View_1 ON SC010300.SC01037 = View_1.SY24002 "
                MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode IS NOT NULL) "
                If MyGroup = "---" And MySupplier = "---" Then
                ElseIf MyGroup <> "---" And MySupplier = "---" Then
                    MySQLStr = MySQLStr & "AND (SC010300.SC01037 = N'" & MyGroup & "') "
                ElseIf MyGroup = "---" And MySupplier <> "---" Then
                    MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & MySupplier & "') "
                Else
                    MySQLStr = MySQLStr & "AND (SC010300.SC01037 = N'" & MyGroup & "') AND (SC010300.SC01058 = N'" & MySupplier & "') "
                End If
            ElseIf MyCodes = 1 Then         '---��������� �������� � ������ ������� ����������
                MySQLStr = "SELECT tbl_WEB_Pictures.SupplierItemCode AS ItemCode, tbl_WEB_Pictures.Picture, Ltrim(Rtrim(ISNULL(View_1.SY24002, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(View_1.SY24003, N'����������'))) "
                MySQLStr = MySQLStr & "AS ProductGroup, Ltrim(Rtrim(ISNULL(PL010300.PL01001, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(PL010300.PL01002, N'����������'))) AS Supplier "
                MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SY24002, SY24003 "
                MySQLStr = MySQLStr & "FROM SY240300 "
                MySQLStr = MySQLStr & "WHERE (SY24001 = N'IB') AND (SY24002 <> N'')) AS View_1 ON SC010300.SC01037 = View_1.SY24002 "
                If MyGroup = "---" And MySupplier = "---" Then
                ElseIf MyGroup <> "---" And MySupplier = "---" Then
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01037 = N'" & MyGroup & "') "
                ElseIf MyGroup = "---" And MySupplier <> "---" Then
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01058 = N'" & MySupplier & "') "
                Else
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01037 = N'" & MyGroup & "') AND (SC010300.SC01058 = N'" & MySupplier & "') "
                End If
            End If
        ElseIf MyRange = 1 Then         '---������ ������ �� ������� ���� ����� ���� �� ������� (��� �� tbl_WEB_Items - ��� �������� �� ����)
            If MyCodes = 0 Then         '---��������� �������� � ������ Scala
                MySQLStr = "SELECT tbl_WEB_Pictures.ScalaItemCode AS ItemCode, tbl_WEB_Pictures.Picture, Ltrim(Rtrim(ISNULL(View_1.SY24002, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(View_1.SY24003, "
                MySQLStr = MySQLStr & "N'����������'))) AS ProductGroup, Ltrim(Rtrim(ISNULL(PL010300.PL01001, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(PL010300.PL01002, N'����������'))) AS Supplier "
                MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Items ON SC010300.SC01001 = tbl_WEB_Items.Code LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SY24002, SY24003 "
                MySQLStr = MySQLStr & "FROM SY240300 "
                MySQLStr = MySQLStr & "WHERE (SY24001 = N'IB') AND (SY24002 <> N'')) AS View_1 ON SC010300.SC01037 = View_1.SY24002 "
                MySQLStr = MySQLStr & "WHERE (tbl_WEB_Pictures.ScalaItemCode IS NOT NULL) "
                If MyGroup = "---" And MySupplier = "---" Then
                ElseIf MyGroup <> "---" And MySupplier = "---" Then
                    MySQLStr = MySQLStr & "AND (SC010300.SC01037 = N'" & MyGroup & "') "
                ElseIf MyGroup = "---" And MySupplier <> "---" Then
                    MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & MySupplier & "') "
                Else
                    MySQLStr = MySQLStr & "AND (SC010300.SC01037 = N'" & MyGroup & "') AND (SC010300.SC01058 = N'" & MySupplier & "') "
                End If
            ElseIf MyCodes = 1 Then         '---��������� �������� � ������ ������� ����������
                MySQLStr = "SELECT tbl_WEB_Pictures.SupplierItemCode AS ItemCode, tbl_WEB_Pictures.Picture, Ltrim(Rtrim(ISNULL(View_1.SY24002, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(View_1.SY24003, N'����������'))) "
                MySQLStr = MySQLStr & "AS ProductGroup, Ltrim(Rtrim(ISNULL(PL010300.PL01001, N'---'))) + ' ' + Ltrim(Rtrim(ISNULL(PL010300.PL01002, N'����������'))) AS Supplier "
                MySQLStr = MySQLStr & "FROM SC010300 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Items ON SC010300.SC01001 = tbl_WEB_Items.Code INNER JOIN "
                MySQLStr = MySQLStr & "tbl_WEB_Pictures ON SC010300.SC01001 = tbl_WEB_Pictures.ScalaItemCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "(SELECT SY24002, SY24003 "
                MySQLStr = MySQLStr & "FROM SY240300 "
                MySQLStr = MySQLStr & "WHERE (SY24001 = N'IB') AND (SY24002 <> N'')) AS View_1 ON SC010300.SC01037 = View_1.SY24002 "
                If MyGroup = "---" And MySupplier = "---" Then
                ElseIf MyGroup <> "---" And MySupplier = "---" Then
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01037 = N'" & MyGroup & "') "
                ElseIf MyGroup = "---" And MySupplier <> "---" Then
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01058 = N'" & MySupplier & "') "
                Else
                    MySQLStr = MySQLStr & "WHERE (SC010300.SC01037 = N'" & MyGroup & "') AND (SC010300.SC01058 = N'" & MySupplier & "') "
                End If
            End If
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("���� �� ������� �� ����� �������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Declarations.MyRec.MoveLast()
            Label5.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()

            i = 0
            While Not Declarations.MyRec.EOF = True
                DownloadOnePicture(MyCatalog, Declarations.MyRec.Fields("ItemCode").Value, Declarations.MyRec.Fields("Picture").Value, Declarations.MyRec.Fields("Supplier").Value, _
                    Declarations.MyRec.Fields("ProductGroup").Value, MyGroupDir, MySupplCatCreated, MyGroupCatCreated)
                i = i + 1
                Label6.Text = i
                Application.DoEvents()
                Declarations.MyRec.MoveNext()
            End While
        End If
    End Sub

    Private Sub DownloadOnePicture(ByVal MyCatalog As String, ByVal MyPictureName As String, ByVal MyPictureByte As Byte(), ByVal MySuppName As String, ByVal MyGroupName As String, _
    ByVal MyGroupDir As Integer, ByVal MySupplCatCreated As Integer, ByVal MyGroupCatCreated As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ����� �������� � �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim stream As New IO.MemoryStream(MyPictureByte)
        Dim picture As Image
        Dim SuppDir As String
        Dim GroupDir As String

        Try
            If MySupplCatCreated = 0 Then       '---����������� ��� ����������� �� �������
                SuppDir = ""
            Else                                '---����������� ��� ����������� �������
                SuppDir = RemoveIllegalChars(MySuppName) + "\"
            End If

            If MyGroupCatCreated = 0 Then       '---����������� ��� ����� ������� �� �������
                GroupDir = ""
            Else                                '---����������� ��� ����� ������� �������
                GroupDir = RemoveIllegalChars(MyGroupName) + "\"
            End If

            picture = Image.FromStream(stream)
            If MyGroupDir = 0 Then          '����������� ���������� -> ������ �������
                If Directory.Exists(MyCatalog + "\" + SuppDir + GroupDir) = False Then
                    Directory.CreateDirectory(MyCatalog + "\" + SuppDir + GroupDir)
                End If
                picture.Save(MyCatalog + "\" + SuppDir + GroupDir + MyPictureName + ".jpg")

            Else                            '����������� ������ ������� -> ����������
                If Directory.Exists(MyCatalog + "\" + GroupDir + SuppDir) = False Then
                    Directory.CreateDirectory(MyCatalog + "\" + GroupDir + SuppDir)
                End If
                picture.Save(MyCatalog + "\" + GroupDir + SuppDir + MyPictureName + ".jpg")
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
        End Try
    End Sub

    Private Function RemoveIllegalChars(ByVal MyStr As String) As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� "�����������" �������� �� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim IllegalStr As String = "\,/,:,*,?," & Chr(34) & ",<,>,|"
        Dim IllegalChr() As String = Split(IllegalStr, ",")

        For i As Integer = 0 To IllegalChr.Length - 1
            MyStr = Replace(MyStr, IllegalChr(i), "")
        Next

        Return MyStr
    End Function
End Class