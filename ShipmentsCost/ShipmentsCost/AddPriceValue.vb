Public Class AddPriceValue

    Private Sub AddPriceValue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

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

        MySQLStr = "SELECT ID, CONVERT(nvarchar, ID) + ' ' + Name AS Name "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_PriceType WITH (NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ID"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "ID"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then

            Declarations.Destination = TextBox3.Text
            Declarations.PriceType = ComboBox1.SelectedValue
            Declarations.PriceFrom = TextBox4.Text
            Declarations.PriceTo = TextBox5.Text
            Declarations.PriceVal = TextBox7.Text
            Declarations.MinCost = TextBox8.Text
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

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� "������� �" �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""������� �..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("� ���� ""������� �..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
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
        '// ��������, ��� � ���� "��" �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""�� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("� ���� ""�� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
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

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���� ""����� ����������"" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox3.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("���� ""������� � ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox4.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("���� ""�� ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox5.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox7.Text) = "" Then
            MsgBox("���� ""���� ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox7.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox8.Text) = "" Then
            MsgBox("���� ""���. ���� ..."" ������ ���� ���������", MsgBoxStyle.Critical, "��������")
            TextBox8.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = 1 And Trim(TextBox3.Text) <> "������� �� �������" Then
            MsgBox("���� �� ������� ��� ����� - ����� ""�� 100 ���������� �� �������"", �� � ���� ����� ���������� ������ ���� �������� ""������� �� �������"".", MsgBoxStyle.Critical, "��������")
            TextBox3.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = 0 And Trim(TextBox3.Text) = "������� �� �������" Then
            MsgBox("���� �� ������� ����� ���������� ""������� �� �������"", �� � ���� ��� ����� - ����� ������ ���� �������� ""�� 100 ���������� �� �������"".", MsgBoxStyle.Critical, "��������")
            ComboBox1.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function

    Private Sub TextBox8_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox8.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������, ��� � ���� "���. ����" �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox8.Text) <> "" Then
            If InStr(TextBox8.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("� ���� ""���. ���� ..."" ������ ���� ������� �����", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox8.Text
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