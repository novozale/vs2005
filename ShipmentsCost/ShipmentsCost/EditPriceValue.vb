Public Class EditPriceValue

    Private Sub EditPriceValue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox6.Text = MainForm.ComboBox1.Text
        TextBox1.Text = MainForm.ComboBox2.Text
        TextBox2.Text = MainForm.ComboBox3.Text
        If MainForm.ComboBox3.SelectedValue = 1 Then    '---����������
            Label5.Text = "������� � ���� "
            Label6.Text = "�� ��� "
            Label7.Text = "���� �� �� (���) "
        Else                                            '---���������
            Label5.Text = "������� � ������ "
            Label6.Text = "�� ����� "
            Label7.Text = "���� �� ��� � (���) "
        End If

        TextBox3.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        TextBox8.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        TextBox4.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
        TextBox5.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString())
        TextBox7.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
        TextBox8.Text = Trim(MainForm.DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString())
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then

            Declarations.PriceVal = TextBox7.Text
            Declarations.MySuccess = True                               '���������� ���������� ��������
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��� ���������� �������� ����� - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub TextBox7_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox7.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� "����" �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox7.Text) <> "" Then
            If InStr(TextBox7.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""���� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox7.Text
                Catch ex As Exception
                    MsgBox("� ���� ""���� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox7.Text) = "" Then
            MsgBox("���� ""���� ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox7.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox9.Text) = "" Then
            MsgBox("���� ""���. ���� ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox9.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function

    Private Sub TextBox9_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox9.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� "���. ����" �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox9.Text) <> "" Then
            If InStr(TextBox9.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""���. ���� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox9.Text
                Catch ex As Exception
                    MsgBox("� ���� ""���. ���� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub
End Class