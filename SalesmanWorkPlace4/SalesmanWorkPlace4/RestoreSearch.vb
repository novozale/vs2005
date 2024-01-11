Public Class RestoreSearch

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyRez1 = 0
        Me.Close()
    End Sub

    Private Sub RestoreSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----���������� �����������
            If SaveRequest() = True Then
                Declarations.MyRez1 = 1
                Me.Close()
            End If
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DateTimePicker1.Value < Now().AddDays(-1) Then
            MsgBox("���� ""���� ������������� ��"" ������ ���� ������ ��� ����� �������", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If
        CheckFormFilling = True

    End Function

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ ��������� � �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        Try
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr & "SET RequestDate = CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103), "
            MySQLStr = MySQLStr & "Comments = Comments + '" + CStr(Chr(10) + Chr(13)) + "' + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & Trim(TextBox1.Text) & "' "
            MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyRequestNum & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        Catch ex As Exception
            MsgBox(ex.ToString)
            SaveRequest = False
            Exit Function
        End Try
        SaveRequest = True
    End Function

    Private Sub RestoreSearch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        MySQLStr = "SELECT RequestDate "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch "
        MySQLStr = MySQLStr & "WHERE (ID = " & Declarations.MyRequestNum & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            DateTimePicker1.Value = Declarations.MyRec.Fields("RequestDate").Value.ToString
        End If
        trycloseMyRec()
    End Sub
End Class