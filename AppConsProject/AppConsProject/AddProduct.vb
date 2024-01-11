Public Class AddProduct

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Declarations.IsSuccess = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        '���������� ������ � �������
        If CheckDataFiling() = True Then
            Declarations.MinQty = TextBox1.Text
            Declarations.MaxQty = TextBox2.Text
            Declarations.IsSuccess = True
            Me.Close()
        End If

    End Sub

    Private Function CheckDataFiling() As Boolean
        
        '// �������� ���������� ����� � ����

        If Trim(TextBox2.Text) = "" Then
            MsgBox("���� ""����������� �������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""������������ �������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function
    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating

        Dim Rez As Double

        If Trim(TextBox2.Text) <> "" Then
            If InStr(TextBox2.Text, System.Globalization.NumberFormatInfo.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""����������� �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    Rez = TextBox2.Text
                Catch ex As Exception
                    MsgBox("� ���� ""������������ �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating

        Dim Rez As Double

        If Trim(TextBox1.Text) <> "" Then
            If InStr(TextBox1.Text, System.Globalization.NumberFormatInfo.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""����������� �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    Rez = TextBox1.Text
                Catch ex As Exception
                    MsgBox("� ���� ""������������ �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub
    
    
End Class