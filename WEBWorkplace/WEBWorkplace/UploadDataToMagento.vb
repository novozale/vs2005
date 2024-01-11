Public Class UploadDataToMagento
    Public MyMode   '--- 0 ������ ��������
    '--- 1 ������ ����� ����������

    Private Sub UploadDataToMagento_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ��� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        Button2.Enabled = False
        Button3.Enabled = False
        GroupBox1.BackColor = Color.LightGray
        GroupBox5.BackColor = Color.LightGray
        GroupBox2.BackColor = Color.LightGray
        GroupBox3.BackColor = Color.LightGray
        GroupBox4.BackColor = Color.LightGray
        GroupBox6.BackColor = Color.LightGray
        Label2.Text = "0"
        Label3.Text = "0"
        Label15.Text = "0"
        Label14.Text = "0"
        Label6.Text = "0"
        Label5.Text = "0"
        Label9.Text = "0"
        Label8.Text = "0"
        Label12.Text = "0"
        Label11.Text = "0"
        Label18.Text = "0"
        Label17.Text = "0"
        If MyMode = 0 Then          '---������ ��������
            UploadInfo_ToMagento(0)
        ElseIf MyMode = 1 Then      '---������ ����� ������
            UploadInfo_ToMagento(1)
        End If
        Me.Cursor = Cursors.Default
        Button2.Enabled = True
        Button3.Enabled = True
        'MsgBox("�������� ������ �� WEB ���� Magento ����������� �������.", MsgBoxStyle.Information, "��������!")
        MyErrWindow = New ErrWindow
        MyErrWindow.ShowDialog()
    End Sub

    Private Sub UploadDataToMagento_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyMode = 0 Then
            Label19.Visible = True
        Else
            Label19.Visible = False
        End If
    End Sub
End Class