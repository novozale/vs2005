Public Class EditCustom

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��� ���������� ������ �������� ���, ROP � ���������� ������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub EditCustom_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Label3.Text = "�������������� ������ �������� ���, ROP � ���������� ������ ��� ������ " & Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & " " & Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString())
        TextBox6.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(8).Value.ToString())
        TextBox4.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(12).Value.ToString())
        TextBox3.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(16).Value.ToString())
        TextBox2.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(6).Value.ToString())
        TextBox1.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(10).Value.ToString())
        TextBox5.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(14).Value.ToString())
    End Sub

    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� ��� �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox2.Text) <> "" Then
            If InStr(TextBox2.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""���"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox2.Text
                Catch ex As Exception
                    MsgBox("� ���� ""���"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� ROP �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox1.Text) <> "" Then
            If InStr(TextBox1.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""ROP"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox1.Text
                Catch ex As Exception
                    MsgBox("� ���� ""ROP"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� ������� ���������� ������ �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""��������� �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("� ���� ""��������� �������"" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ROP � ���������� ������ ��� ������� ��� ��� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) <> "" Then
            If CDbl(TextBox6.Text) <> 0 Then
                TextBox1.Text = Math.Round(CDbl(TextBox2.Text) * CDbl(TextBox4.Text) / CDbl(TextBox6.Text), 3)
                TextBox5.Text = Math.Round(CDbl(TextBox2.Text) * CDbl(TextBox3.Text) / CDbl(TextBox6.Text), 3)
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then

            Declarations.MyMGZ = TextBox2.Text
            Declarations.MyROP = TextBox1.Text
            Declarations.MyInsuranceLVL = TextBox5.Text
            Declarations.MySuccess = True                               '���������� ���������� ��������
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
            MsgBox("���� ""���"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""ROP"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("���� ""��������� �������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function
End Class