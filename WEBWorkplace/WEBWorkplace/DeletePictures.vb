Public Class DeletePictures

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
            MsgBox("������� � ���������� ��� �������� � �������� ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub DeletePictures_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ������ �� alt - F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
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
        '// �������� � �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            ComboBox1.Enabled = False
            CheckAndDeletePictures(Trim(TextBox1.Text), ComboBox1.SelectedValue)
            MsgBox("�������� � �������� �������� ���������.", MsgBoxStyle.Information, "��������!")
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            ComboBox1.Enabled = True
        End If

    End Sub

    Private Sub DeletePictures_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ����� �������
        Dim MyDs As New DataSet

        InitMyConn(False)
        '---���������� �������� � �������
        MySQLStr = "SELECT DISTINCT ID, Convert(nvarchar(10),ID) + '  ' + Ltrim(Rtrim(CompanyName)) AS CompanyName "
        MySQLStr = MySQLStr & "FROM tbl_WEBSearchScrapping_Companies "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "CompanyName" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CheckAndDeletePictures(ByVal MyCatalog As String, ByVal MyPictSupplier As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer
        Dim j As Integer

        MySQLStr = "SELECT DISTINCT Ltrim(Rtrim(SC010300.SC01060)) as PictureName "
        MySQLStr = MySQLStr & "FROM tbl_WEBSearchScrapping_NotCorrectSuppl INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_WEBSearchScrapping_NotCorrectSuppl.PL01001 = SC010300.SC01058 "
        MySQLStr = MySQLStr & "WHERE(tbl_WEBSearchScrapping_NotCorrectSuppl.CompanyID = " & MyPictSupplier & ") "
        MySQLStr = MySQLStr & "ORDER BY Ltrim(Rtrim(SC010300.SC01060)) "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("��� ������� ���������� �������� �������� � �������� �� �����.", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        Else
            Declarations.MyRec.MoveLast()
            Label5.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()

            i = 0
            j = 0
            While Not Declarations.MyRec.EOF = True
                j = j + CheckAndDeleteOnePicture(MyCatalog, Declarations.MyRec.Fields("PictureName").Value)
                i = i + 1
                Label6.Text = i
                Label8.Text = j
                Application.DoEvents()
                Declarations.MyRec.MoveNext()
            End While
        End If
    End Sub

    Private Function CheckAndDeleteOnePicture(ByVal MyCatalog As String, ByVal MyPictName As String) As Integer
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � �������� ����� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Try
            'My.Computer.FileSystem.DeleteFile(MyCatalog & "\" & MyPictName & ".*", Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
            Kill(MyCatalog & "\" & MyPictName & ".*")
            CheckAndDeleteOnePicture = 1
        Catch ex As Exception
            CheckAndDeleteOnePicture = 0
        End Try
    End Function
End Class