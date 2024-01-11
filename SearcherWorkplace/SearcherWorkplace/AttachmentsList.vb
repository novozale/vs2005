Public Class AttachmentsList
    Public AttType As String
    Public WhoStart As String
    Public MyPlace As String

    Private Sub AttachmentsList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub AttachmentsList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных об аттачментах в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////

        loadAtttList()
    End Sub

    Private Sub loadAtttList()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных об аттачментах в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyAdapter4 As SqlClient.SqlDataAdapter    'для списка аттачментов
        Dim MyDs4 As New DataSet
        Dim MySQLStr As String

        MySQLStr = "SELECT ID, SupplSearchID, AttachmentName "
        If AttType = "Sales" Then
            MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SalesAttachments WITH(NOLOCK) "
        Else
            MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SearchAttachments WITH(NOLOCK) "
        End If
        MySQLStr = MySQLStr & "WHERE (SupplSearchID = '" & Declarations.MyRequestNum & "') "
        MySQLStr = MySQLStr & "ORDER BY AttachmentName "
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.SelectCommand.CommandTimeout = 600
            MyAdapter4.Fill(MyDs4)
            DataGridView1.DataSource = MyDs4.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 0
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "ID запроса"
        DataGridView1.Columns(1).Width = 0
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "Имя файла"
        DataGridView1.Columns(2).Width = 550

        '----------права доступа
        If DataGridView1.SelectedRows.Count = 0 Then
            If AttType = WhoStart Then
                Button6.Enabled = True
            Else
                Button6.Enabled = False
            End If
            Button7.Enabled = False
            Button8.Enabled = False
        Else
            If AttType = WhoStart Then
                Button6.Enabled = True
                Button7.Enabled = True
            Else
                Button6.Enabled = False
                Button7.Enabled = False
            End If
            Button8.Enabled = True
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// занесение нового аттачмента в БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim FName As String
        Dim i As Integer
        Dim mstream As ADODB.Stream
        Dim FInfo As FileInfo
        Dim AttID As Integer

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If (OpenFileDialog1.FileName <> "") Then
                '--имя файла без пути
                i = InStrRev(OpenFileDialog1.FileName, "\")
                FName = Microsoft.VisualBasic.Right(OpenFileDialog1.FileName, Len(OpenFileDialog1.FileName) - i)
                FInfo = New FileInfo(OpenFileDialog1.FileName)
                If FInfo.Length <> 0 Then
                    Try
                        AttID = 0
                        MySQLStr = "exec spp_SupplSearch_AddAttachment "
                        MySQLStr = MySQLStr & Declarations.MyRequestNum.ToString & ", "
                        MySQLStr = MySQLStr & "N'" & FName & "', "
                        If AttType = "Sales" Then
                            MySQLStr = MySQLStr & "N'Sales' "
                        Else
                            MySQLStr = MySQLStr & "N'Search' "
                        End If
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            MsgBox("Ошибка создания записи")
                            trycloseMyRec()
                        Else
                            AttID = Declarations.MyRec.Fields("MyNewID").Value
                            trycloseMyRec()
                        End If

                        If AttID <> 0 Then
                            MySQLStr = "SELECT ID, SupplSearchID, AttachmentName, AttachmentBody "
                            If AttType = "Sales" Then
                                MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SalesAttachments "
                            Else
                                MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SearchAttachments "
                            End If
                            MySQLStr = MySQLStr & "WHERE (ID = '" & AttID & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            mstream = New ADODB.Stream
                            mstream.Type = StreamTypeEnum.adTypeBinary
                            mstream.Open()
                            mstream.LoadFromFile(OpenFileDialog1.FileName)
                            Declarations.MyRec.Fields("AttachmentBody").Value = mstream.Read
                            Declarations.MyRec.Update()
                            trycloseMyRec()
                            loadAtttList()
                        End If
                    Catch ex As Exception
                        If AttType = "Sales" Then
                            MySQLStr = "DELETE FROM tbl_SupplSearch_SalesAttachments "
                        Else
                            MySQLStr = "DELETE FROM tbl_SupplSearch_SearchAttachments "
                        End If
                        MySQLStr = MySQLStr & "WHERE (ID = '" & AttID & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                        MsgBox(ex.ToString)
                    End Try
                Else
                    MsgBox("Файл " & OpenFileDialog1.FileName & " имеет нулевой размер и не может быть импортирован. ", MsgBoxStyle.Critical, "Внимние!")
                End If
            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// удаление аттачмента из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If AttType = "Sales" Then
            MySQLStr = "DELETE FROM tbl_SupplSearch_SalesAttachments "
        Else
            MySQLStr = "DELETE FROM tbl_SupplSearch_SearchAttachments "
        End If
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        loadAtttList()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Извлечение аттачмента из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim mstream As ADODB.Stream

        Try
            MySQLStr = "SELECT ID, SupplSearchID, AttachmentName, AttachmentBody "
            If AttType = "Sales" Then
                MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SalesAttachments "
            Else
                MySQLStr = MySQLStr & "FROM tbl_SupplSearch_SearchAttachments "
            End If
            MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "ORDER BY AttachmentName "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            SaveFileDialog1.FileName = Declarations.MyRec.Fields("AttachmentName").Value
            If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If (SaveFileDialog1.FileName <> "") Then
                    mstream = New ADODB.Stream
                    mstream.Type = StreamTypeEnum.adTypeBinary
                    mstream.Open()
                    mstream.Write(Declarations.MyRec.Fields("AttachmentBody").Value)
                    mstream.SaveToFile(SaveFileDialog1.FileName, SaveOptionsEnum.adSaveCreateOverWrite)
                End If
            End If
            trycloseMyRec()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyPlace = "List" Then
            If AttType = "Sales" And WhoStart = "Sales" Then
                MainForm.DataGridView1.SelectedRows.Item(0).Cells(15).Value = DataGridView1.RowCount
            End If
            If AttType = "Search" And WhoStart = "Search" Then
                MainForm.DataGridView1.SelectedRows.Item(0).Cells(16).Value = DataGridView1.RowCount
            End If
        End If
        Me.Close()
    End Sub
End Class