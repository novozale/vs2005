Imports System.IO

Public Class LoadOnePictToDB

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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
        '// ����� �������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetPictureFile()
        If MyCatalog = "" Then      '--������ ������
        Else
            TextBox3.Text = MyCatalog
            PictureBox1.Image = Image.FromFile(MyCatalog)
        End If
    End Sub

    Private Function GetPictureFile() As String
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� - ����� �������� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        OpenFileDialog1.ShowDialog()
        GetPictureFile = OpenFileDialog1.FileName
    End Function

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ���� �� ������ ��� � Scala
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) <> "" Then
            MySQLStr = "SELECT COUNT(SC01001) AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(TextBox1.Text) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("��������� ��� ������ " & Trim(TextBox1.Text) & " �� ������ � Scala ")
                e.Cancel = True
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    MsgBox("��������� ��� ������ " & Trim(TextBox1.Text) & " �� ������ � Scala ")
                    e.Cancel = True
                Else

                End If
            End If
        End If
    End Sub

    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� - ���� �� ������ ��� ������ ���������� � Scala
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox2.Text) <> "" Then
            MySQLStr = "SELECT COUNT(SC01060) AS CC "
            MySQLStr = MySQLStr & "FROM SC010300 "
            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(TextBox2.Text) & "') AND (SC01060 <> '') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("��������� ��� ������ ���������� " & Trim(TextBox2.Text) & " �� ������ � Scala ")
                e.Cancel = True
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    MsgBox("��������� ��� ������ ����������" & Trim(TextBox2.Text) & " �� ������ � Scala ")
                    e.Cancel = True
                Else
                    trycloseMyRec()
                End If
            End If
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLstr As String

        If Trim(TextBox1.Text) = "" Then
            MsgBox("��� �������� � Scala ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("��� ������ ���������� � Scala ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            TextBox2.Select()
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���� � ��������� ����������� ������ ���� ������.", MsgBoxStyle.Critical, "��������!")
            CheckData = False
            Button1.Select()
            Exit Function
        End If

        MySQLstr = "SELECT COUNT(SC01001) AS CC "
        MySQLstr = MySQLstr & "FROM SC010300 "
        MySQLstr = MySQLstr & "WHERE (SC01060 = N'" & Trim(TextBox2.Text) & "') "
        MySQLstr = MySQLstr & "AND (SC01001 = N'" & Trim(TextBox1.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLstr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("������ � ����� " & Trim(TextBox1.Text) & " � ����� ������ ���������� " & Trim(TextBox2.Text) & " ��� � Scala ")
            CheckData = False
            TextBox1.Select()
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                trycloseMyRec()
                MsgBox("������ � ����� " & Trim(TextBox1.Text) & " � ����� ������ ���������� " & Trim(TextBox2.Text) & " ��� � Scala ")
                CheckData = False
                TextBox1.Select()
                Exit Function
            Else
                trycloseMyRec()
            End If
        End If

        CheckData = True
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������������� �������� � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ExistFlag As Integer
        Dim MyAnswer As Object

        If CheckData() = True Then
            ExistFlag = 1

            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
            MySQLStr = MySQLStr & "WHERE (ScalaItemCode = N'" & Trim(TextBox1.Text) & "') "
            MySQLStr = MySQLStr & "AND (SupplierItemCode = N'" & Trim(TextBox2.Text) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                ExistFlag = 0
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    ExistFlag = 0
                Else
                    trycloseMyRec()
                    ExistFlag = 1
                End If
            End If

            MyAnswer = DialogResult.Yes
            If ExistFlag = 1 Then
                MyAnswer = MsgBox("��� ������ � ����� " & Trim(TextBox1.Text) & " � ����� ������ ���������� " & Trim(TextBox2.Text) & " ��� ���� �������� � ��. ������������?", MsgBoxStyle.YesNo, "��������!")
            End If
            If MyAnswer = DialogResult.Yes Then
                UploadMyPicture(Trim(TextBox3.Text), Trim(TextBox1.Text))
                '---�������� ����� ����������
                MySQLStr = "UPDATE tbl_WEB_Items "
                MySQLStr = MySQLStr & "SET RMStatus = CASE WHEN tbl_WEB_Items.RMStatus = 1 THEN 1 ELSE CASE WHEN tbl_WEB_Items.RMStatus = 2 THEN 2 ELSE 3 END END, "
                MySQLStr = MySQLStr & "WEBStatus = CASE WHEN tbl_WEB_Items.WEBStatus = 1 THEN 1 ELSE CASE WHEN tbl_WEB_Items.WEBStatus = 2 THEN 2 ELSE 3 END END "
                MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(TextBox1.Text) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub UploadMyPicture(ByVal MyCatalog As String, ByVal ScalaCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ��������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Dim Finfo As New FileInfo(MyCatalog)
        If Finfo.Length = 0 Then
            MsgBox("���� � ����������� .jpg ����� ������� ������.", MsgBoxStyle.Critical, "��������!")
        Else
            WritePictureToDB(MyCatalog, FileIO.FileSystem.GetName(MyCatalog).Substring(0, Len(FileIO.FileSystem.GetName(MyCatalog)) - 4), ScalaCode)
        End If
    End Sub

    Private Sub WritePictureToDB(ByVal MyPicturePath As String, ByVal MyPictureName As String, ByVal ScalaCode As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������ ����� �������� � ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim connection As SqlClient.SqlConnection
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_Pictures "
        MySQLStr = MySQLStr & "WHERE (ScalaItemCode = N'" & ScalaCode & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        Try
            connection = New SqlClient.SqlConnection(Declarations.MyNETConnStr)

            Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand("INSERT INTO tbl_WEB_Pictures " _
                + "(ID, ScalaItemCode, SupplierItemCode, PictureSmall, PictureMedium, Picture) " _
                + "VALUES(NEWID(), @ScalaItemCode, @SupplierItemCode, @PictureSmallBytes, @PictureMediumBytes, @PictureBytes) ", connection)

            Dim par As SqlClient.SqlParameter = New SqlClient.SqlParameter("@SupplierItemCode", SqlDbType.NVarChar)
            par.Value = MyPictureName.ToString
            par.Direction = ParameterDirection.Input
            cmd.Parameters.Add(par)

            par = New SqlClient.SqlParameter("@ScalaItemCode", SqlDbType.NVarChar)
            par.Value = ScalaCode
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
                imageMediumBytes = UploadFilesToDB.MkTh(imageBytes, 100, 100)
                par.Value = imageMediumBytes
                cmd.Parameters.Add(par)

                par = New SqlClient.SqlParameter("@PictureSmallBytes", SqlDbType.Image)
                par.Direction = ParameterDirection.Input
                Dim imageSmallBytes(1225) As Byte
                imageSmallBytes = UploadFilesToDB.MkTh(imageBytes, 35, 35)
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