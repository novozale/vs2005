Public Class ErrForm
    Public MyErrStr As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ����������� 0 (����������� ���������� ������)  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyErrRezult = 0
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � ����������� 1 (���������� ��������)  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyErrRezult = 1
        Me.Close()
    End Sub

    Private Sub ErrForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        RichTextBox1.Text = MyErrStr
    End Sub
End Class