Public Class UploadInfoToSaintGobain

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

    Private Sub UploadInfoToSaintGobain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        '// �������� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyRange As Integer

        If Not TextBox1.Text.Equals("") Then
            Me.Cursor = Cursors.WaitCursor
            Select Case ComboBox1.SelectedItem
                Case "��� ������"
                    MyRange = 0
                Case "��� ������ � ����������"
                    MyRange = 1
                Case "��� ������ � ���������� � ����������"
                    MyRange = 2
                Case "������������� �����������"
                    MyRange = 3
                Case "������������� ����������� � ����������"
                    MyRange = 4
                Case "������������� ����������� � ���������� � ����������"
                    MyRange = 5
                Case Else
                    MyRange = 0
            End Select

            If My.Settings.UseOffice = "LibreOffice" Then
                UploadSenGobenToLO(TextBox1.Text, MyRange)
            Else
                UploadSenGobenToExcel(TextBox1.Text, MyRange)
            End If

            MsgBox("�������� ���������� ���������", MsgBoxStyle.OkOnly, "��������!")
            Me.Cursor = Cursors.Default
        Else
            MsgBox("���������� ������� ������� ��� �������� ��������", MsgBoxStyle.Critical, "��������!")
            Button1.Select()
        End If
    End Sub

    Private Sub UploadInfoToSaintGobain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ComboBox1.SelectedItem = "������������� ����������� � ���������� � ����������"
    End Sub

End Class